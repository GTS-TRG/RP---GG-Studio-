/**
 * File System Access API — mở file + lưu bản sao an toàn (Chrome, Edge).
 * Mặc định KHÔNG ghi đè file gốc. Khi ghi, luôn validate buffer trước.
 */

export type ExcelFileHandle = FileSystemFileHandle;

const EXCEL_PICKER_TYPES = [
  {
    description: 'Excel workbook',
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'],
      'application/vnd.ms-excel': ['.xls'],
    },
  },
];

export function supportsFileSystemAccess(): boolean {
  return typeof window !== 'undefined' && typeof window.showOpenFilePicker === 'function';
}

export function supportsSaveFilePicker(): boolean {
  return typeof window !== 'undefined' && typeof window.showSaveFilePicker === 'function';
}

/** Mở file Excel kèm handle (để biết path gốc; Save mặc định ghi bản sao mới) */
export async function openExcelWithHandle(): Promise<{
  handle: ExcelFileHandle;
  file: File;
  buffer: ArrayBuffer;
} | null> {
  if (!supportsFileSystemAccess() || !window.showOpenFilePicker) return null;

  try {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      excludeAcceptAllOption: false,
      types: EXCEL_PICKER_TYPES,
    });

    // Xin quyền đọc/ghi (nếu sau này user chủ động chọn Overwrite)
    try {
      const opts: FileSystemHandlePermissionDescriptor = { mode: 'readwrite' };
      const q = await handle.queryPermission(opts);
      if (q !== 'granted') {
        await handle.requestPermission(opts);
      }
    } catch {
      // ignore
    }

    const file = await handle.getFile();
    if (!file || file.size < 64) {
      throw new Error(
        'This Excel file is empty or corrupted (file size too small).\n\n' +
          'Restore the file from backup / OneDrive version history, then open again.\n' +
          'Do not use the damaged file.'
      );
    }

    const buffer = await file.arrayBuffer();
    if (!buffer || buffer.byteLength < 64) {
      throw new Error(
        'Could not read file contents (empty buffer).\nClose the file in Excel and try again.'
      );
    }

    // Kiểm tra header ZIP (xlsx/xlsm)
    const u8 = new Uint8Array(buffer);
    const isZip = u8[0] === 0x50 && u8[1] === 0x4b;
    if (!isZip && !/\.xls$/i.test(file.name)) {
      throw new Error(
        'File does not look like a valid Excel package (missing ZIP header).\n' +
          'The file may already be corrupted. Restore a backup before opening.'
      );
    }

    return { handle, file, buffer };
  } catch (err) {
    if (err instanceof DOMException && err.name === 'AbortError') return null;
    throw err;
  }
}

/**
 * Doc lai file tu handle da mo (Reload) - khong can chon lai file.
 * Dung khi user vua sua file trong Excel va bam Save.
 */
export async function readFileFromHandle(handle: ExcelFileHandle): Promise<{
  file: File;
  buffer: ArrayBuffer;
}> {
  let file: File;
  try {
    file = await handle.getFile();
  } catch (err) {
    const name = err instanceof DOMException ? err.name : '';
    if (name === 'NotAllowedError') {
      throw new Error(
        'Quyen doc file da het hieu luc. Bam "Open workbook" de mo lai file.'
      );
    }
    if (name === 'NotFoundError') {
      throw new Error('File da bi di chuyen hoac xoa. Bam "Open workbook" de mo lai.');
    }
    throw err;
  }

  if (!file || file.size < 64) {
    throw new Error(
      'File tren dia dang rong hoac hong (kich thuoc qua nho). ' +
        'Kiem tra lai file trong Excel truoc khi Reload.'
    );
  }

  const buffer = await file.arrayBuffer();
  if (!buffer || buffer.byteLength < 64) {
    throw new Error(
      'Khong doc duoc noi dung file. Neu Excel dang ghi file, doi vai giay roi Reload lai.'
    );
  }

  const u8 = new Uint8Array(buffer);
  const isZip = u8[0] === 0x50 && u8[1] === 0x4b;
  if (!isZip && !/\.xls$/i.test(file.name)) {
    throw new Error(
      'File khong phai goi Excel hop le (thieu ZIP header). ' +
        'Co the Excel dang luu do dang - doi vai giay roi Reload lai.'
    );
  }

  return { file, buffer };
}

function explainWriteError(err: unknown): string {
  const name = err instanceof DOMException ? err.name : '';
  const msg = err instanceof Error ? err.message : String(err);

  if (name === 'NotAllowedError' || /permission|denied/i.test(msg)) {
    return 'Write permission was denied. Close Excel if the file is open, then try again.';
  }
  if (name === 'NoModificationAllowedError' || /locked|in use|busy|sharing/i.test(msg)) {
    return 'The file is locked (often open in Excel). Close it, then try again.';
  }
  if (name === 'NotFoundError') {
    return 'The file was moved or deleted. Please open it again.';
  }
  if (name === 'AbortError') {
    return 'Save cancelled.';
  }
  return msg || 'Unknown write error';
}

async function ensureReadWrite(handle: ExcelFileHandle): Promise<void> {
  try {
    const opts: FileSystemHandlePermissionDescriptor = { mode: 'readwrite' };
    let perm = await handle.queryPermission(opts);
    if (perm !== 'granted') {
      perm = await handle.requestPermission(opts);
    }
    if (perm !== 'granted') {
      throw new Error('Write permission denied.');
    }
  } catch (err) {
    if (err instanceof Error && err.message === 'Write permission denied.') throw err;
    // Một số browser không có queryPermission — createWritable sẽ kiểm tra
  }
}

/**
 * Ghi buffer đã validate vào handle.
 * - Chỉ ghi khi buffer có header ZIP hợp lệ và size >= 64
 * - Ghi một lần bằng Uint8Array (tránh Blob/seek truncation lỗi)
 * - keepExistingData: false: ghi file mới sạch từ buffer đầy đủ (đã validate trước đó)
 */
export async function writeBufferToHandle(
  handle: ExcelFileHandle,
  buffer: ArrayBuffer
): Promise<void> {
  if (!buffer || buffer.byteLength < 64) {
    throw new Error('Refusing to write: buffer is empty. Original file was not modified.');
  }
  const u8 = new Uint8Array(buffer);
  if (!(u8[0] === 0x50 && u8[1] === 0x4b)) {
    throw new Error('Refusing to write: buffer is not a valid XLSX package. Original file was not modified.');
  }

  try {
    await ensureReadWrite(handle);

    // bytes đầy đủ — copy tách biệt để tránh view lệch
    const bytes = new Uint8Array(buffer.byteLength);
    bytes.set(u8);

    const writable = await handle.createWritable({ keepExistingData: false });
    try {
      await writable.write(bytes);
      await writable.close();
    } catch (writeErr) {
      // Cố đóng stream; file có thể đã truncate — báo rõ
      try {
        await writable.abort();
      } catch {
        /* ignore */
      }
      throw new Error(
        explainWriteError(writeErr) +
          '\n\nWARNING: The target file may be incomplete. Restore from backup if Excel cannot open it.'
      );
    }
  } catch (err) {
    if (err instanceof Error && /Refusing to write|WARNING:/.test(err.message)) throw err;
    throw new Error(explainWriteError(err));
  }
}

/**
 * Lưu AN TOÀN: chọn file mới / Save As (không đụng file gốc trừ khi user chọn đúng path).
 * Trả về tên file đã lưu, hoặc null nếu user hủy.
 */
export async function saveBufferAsNewFile(
  buffer: ArrayBuffer,
  suggestedName: string
): Promise<string | null> {
  if (!supportsSaveFilePicker() || !window.showSaveFilePicker) {
    return null;
  }

  if (!buffer || buffer.byteLength < 64) {
    throw new Error('Refusing to save: empty buffer.');
  }

  try {
    const safeName = suggestedName.toLowerCase().endsWith('.xlsx')
      ? suggestedName
      : suggestedName.toLowerCase().endsWith('.xlsm')
      ? suggestedName
      : `${suggestedName.replace(/\.[^/.]+$/, '')}.xlsx`;

    const handle = await window.showSaveFilePicker({
      suggestedName: safeName,
      types: [
        {
          description: 'Excel workbook',
          accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
            'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'],
          },
        },
      ],
    });

    await writeBufferToHandle(handle, buffer);
    return handle.name;
  } catch (err) {
    if (err instanceof DOMException && err.name === 'AbortError') return null;
    if (err instanceof Error) throw err;
    throw new Error(explainWriteError(err));
  }
}
