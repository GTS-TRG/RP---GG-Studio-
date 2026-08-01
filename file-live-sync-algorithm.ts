/**
 * ============================================================================
 * THUẬT TOÁN THEO DÕI FILE EXCEL TRÊN Ổ ĐĨA THEO THỜI GIAN THỰC (LIVE SYNC)
 * ============================================================================
 * Trích từ dự án "GG Studio - Panel Schedule Reviewer" (src/utils/fileSystemAccess.ts
 * + phần useEffect polling trong src/App.tsx).
 *
 * Ý TƯỞNG:
 * Trình duyệt KHÔNG có API "watch file thay đổi trên ổ đĩa" (không giống Node.js
 * fs.watch). Nhưng với File System Access API (Chrome/Edge), sau khi người dùng
 * chọn 1 file bằng showOpenFilePicker(), web app giữ được 1 "FileSystemFileHandle"
 * trỏ thẳng tới file đó trên đĩa (không phải bản copy trong bộ nhớ trình duyệt).
 *
 * Từ handle đó, bất cứ lúc nào gọi handle.getFile() sẽ trả về File mới nhất,
 * kèm 2 trường rẻ để kiểm tra thay đổi mà KHÔNG cần đọc toàn bộ nội dung:
 *   - file.lastModified (timestamp sửa đổi cuối)
 *   - file.size          (kích thước byte)
 *
 * => Thuật toán: cứ mỗi N mili-giây (vd 1500ms), gọi getFile() lấy 2 giá trị này,
 *    so với "dấu vân tay" (fingerprint) đã lưu lần đọc trước:
 *      - Không đổi -> bỏ qua, không làm gì (rẻ, không tốn băng thông đọc file).
 *      - Có đổi (khác mtime HOẶC khác size) -> nghĩa là người dùng vừa Ctrl+S
 *        trong Excel -> mới đọc lại toàn bộ nội dung (arrayBuffer) và xử lý.
 *
 * VẤN ĐỀ CẦN XỬ LÝ:
 * 1. Excel không ghi file trong 1 thao tác duy nhất — có thể ghi file tạm rồi
 *    rename, hoặc ghi nhiều lần liên tiếp trong vài trăm ms. Nếu đọc ngay lúc
 *    đang ghi dở sẽ ra file rỗng/hỏng.
 *    -> Giải pháp: khi phát hiện đổi, đợi thêm ~400ms rồi getFile() lại lần 2,
 *       so sánh (mtime, size) của 2 lần đọc liên tiếp — CHỈ xử lý khi 2 lần
 *       đọc cho kết quả GIỐNG NHAU (nghĩa là file đã "ổn định", ghi xong).
 *       Nếu vẫn khác nhau -> bỏ qua, để nhịp poll sau xử lý tiếp.
 * 2. Không được đọc chồng lấp (2 lần xử lý cùng lúc) -> dùng 1 cờ "busy" (ref/biến),
 *    tick nào thấy đang busy thì bỏ qua luôn.
 * 3. Nếu người dùng đang có thay đổi CHƯA LƯU ngay trên web app (vd sửa 1 ô trên
 *    UI nhưng chưa ghi xuống đĩa) thì KHÔNG được tự động ghi đè state hiện tại —
 *    chỉ nên bật cờ "diskChanged" để hỏi người dùng trước khi tải lại.
 * 4. Quyền truy cập file (permission) có thể hết hạn / bị thu hồi giữa chừng
 *    (đổi tab lâu, trình duyệt tự thu hồi...) -> bọc try/catch quanh getFile(),
 *    lỗi thì bỏ qua nhịp đó, thử lại nhịp sau (không throw làm crash app).
 *
 * File này KHÔNG phụ thuộc React lẫn phần còn lại của dự án gốc — phần lõi
 * (FileHandleWatcher) là vanilla TypeScript, dùng được cho bất kỳ project nào
 * chạy trên trình duyệt hỗ trợ File System Access API (Chrome, Edge; Safari/
 * Firefox hiện chưa hỗ trợ — cần kiểm tra supportsFileSystemAccess() trước).
 * Phần cuối file có 1 React hook mỏng bọc quanh class này (tuỳ chọn dùng).
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// 0. AMBIENT TYPES cho File System Access API — 1 vài trình duyệt/TypeScript lib
//    chưa có sẵn type này. Nếu project khác báo lỗi "Cannot find name
//    FileSystemHandlePermissionDescriptor", giữ nguyên khai báo dưới đây;
//    nếu project đã có type riêng (vd cài @types/wicg-file-system-access) thì
//    xoá khối này để tránh khai báo trùng.
// ----------------------------------------------------------------------------

declare global {
  interface FileSystemHandlePermissionDescriptor {
    mode?: 'read' | 'readwrite';
  }
}

// ----------------------------------------------------------------------------
// 1. KIỂM TRA HỖ TRỢ TRÌNH DUYỆT
// ----------------------------------------------------------------------------

export function supportsFileSystemAccess(): boolean {
  return typeof window !== 'undefined' && typeof (window as any).showOpenFilePicker === 'function';
}

export function supportsSaveFilePicker(): boolean {
  return typeof window !== 'undefined' && typeof (window as any).showSaveFilePicker === 'function';
}

// ----------------------------------------------------------------------------
// 2. MỞ FILE + LẤY HANDLE (giữ liên kết tới file thật trên đĩa)
// ----------------------------------------------------------------------------

export interface OpenFileResult {
  handle: FileSystemFileHandle;
  file: File;
  buffer: ArrayBuffer;
}

export interface FilePickerAcceptType {
  description: string;
  accept: Record<string, string[]>;
}

/**
 * Mở hộp thoại chọn file, xin quyền đọc/ghi, trả về handle + nội dung file (buffer).
 * `acceptTypes` tuỳ biến theo loại file của dự án khác (mặc định: mọi file).
 * `minValidSize` để chặn sớm file rỗng/hỏng (tuỳ chọn, ví dụ .xlsx tối thiểu 64 bytes).
 */
export async function openFileWithHandle(
  acceptTypes?: FilePickerAcceptType[],
  minValidSize = 0
): Promise<OpenFileResult | null> {
  if (!supportsFileSystemAccess()) return null;

  try {
    const [handle] = await (window as any).showOpenFilePicker({
      multiple: false,
      excludeAcceptAllOption: false,
      types: acceptTypes,
    });

    // Xin quyền đọc/ghi ngay khi mở — để lần sau ghi lại không phải hỏi lại
    try {
      const opts: FileSystemHandlePermissionDescriptor = { mode: 'readwrite' };
      const q = await handle.queryPermission(opts);
      if (q !== 'granted') {
        await handle.requestPermission(opts);
      }
    } catch {
      // Một số trình duyệt/handle không hỗ trợ query/requestPermission — bỏ qua
    }

    const file = await handle.getFile();
    if (!file || file.size < minValidSize) {
      throw new Error('File rỗng hoặc quá nhỏ so với ngưỡng hợp lệ.');
    }

    const buffer = await file.arrayBuffer();
    return { handle, file, buffer };
  } catch (err) {
    if (err instanceof DOMException && err.name === 'AbortError') return null; // user bấm Cancel
    throw err;
  }
}

/** Đọc lại nội dung mới nhất từ 1 handle đã mở trước đó (không cần mở lại hộp thoại) */
export async function readFileFromHandle(
  handle: FileSystemFileHandle
): Promise<{ file: File; buffer: ArrayBuffer }> {
  const file = await handle.getFile();
  const buffer = await file.arrayBuffer();
  return { file, buffer };
}

// ----------------------------------------------------------------------------
// 3. FINGERPRINT — so sánh rẻ (không đọc nội dung) để biết file có đổi không
// ----------------------------------------------------------------------------

export interface FileStamp {
  mtime: number; // file.lastModified
  size: number; // file.size
}

export function stampOf(file: File): FileStamp {
  return { mtime: file.lastModified, size: file.size };
}

export function stampChanged(a: FileStamp | null, b: FileStamp): boolean {
  return !a || a.mtime !== b.mtime || a.size !== b.size;
}

// ----------------------------------------------------------------------------
// 4. FileHandleWatcher — vòng lặp polling lõi, không phụ thuộc framework
// ----------------------------------------------------------------------------

export interface FileHandleWatcherOptions {
  /** Chu kỳ kiểm tra (ms). Mặc định 1500ms — đủ nhanh để cảm giác "tức thời", đủ nhẹ để không tốn CPU. */
  pollMs?: number;
  /** Thời gian chờ file "ổn định" sau khi phát hiện đổi, trước khi đọc lại lần 2 để xác nhận (ms). */
  settleDelayMs?: number;
  /** Gọi khi phát hiện file đã đổi VÀ đã ổn định (ghi xong) — nơi để đọc buffer & xử lý lại logic của app. */
  onChanged: (info: { file: File; stamp: FileStamp }) => void | Promise<void>;
  /**
   * Gọi mỗi tick TRƯỚC khi quyết định đọc lại — trả về true để HOÃN xử lý
   * (vd: người dùng đang có thay đổi chưa lưu trên UI, không muốn bị ghi đè).
   * Nếu trả true, watcher chỉ cập nhật cờ "có thay đổi đang chờ" qua onPendingChange
   * mà KHÔNG gọi onChanged.
   */
  shouldDefer?: () => boolean;
  /** Gọi khi phát hiện đổi nhưng đang bị hoãn (shouldDefer() === true) */
  onPendingChange?: () => void;
  /** Gọi khi getFile()/đọc lỗi (quyền bị thu hồi, file bị khoá...) — không throw, chỉ log/callback */
  onError?: (err: unknown) => void;
}

/**
 * Watcher polling 1 FileSystemFileHandle, tự xử lý:
 *  - So sánh fingerprint (mtime/size) mỗi tick, không đọc nội dung nếu chưa đổi.
 *  - Debounce "settle": đợi rồi đọc lại lần 2, chỉ báo đổi khi 2 lần đọc khớp nhau
 *    (file đã ghi xong, không phải đang ghi dở).
 *  - Chặn xử lý chồng lấp bằng cờ busy nội bộ.
 *  - Không tự ý ghi đè nếu shouldDefer() trả true (có thay đổi chưa lưu trên app).
 *
 * Dùng cho bất kỳ project nào (không chỉ Excel) cần "theo dõi 1 file trên đĩa
 * và tự cập nhật UI ngay khi người dùng Ctrl+S ở app khác (Excel, VSCode, v.v.)".
 */
export class FileHandleWatcher {
  private handle: FileSystemFileHandle;
  private options: Required<Omit<FileHandleWatcherOptions, 'shouldDefer' | 'onPendingChange' | 'onError'>> &
    Pick<FileHandleWatcherOptions, 'shouldDefer' | 'onPendingChange' | 'onError'>;
  private timerId: number | null = null;
  private busy = false;
  private lastStamp: FileStamp | null = null;
  private stopped = true;

  constructor(handle: FileSystemFileHandle, options: FileHandleWatcherOptions, initialStamp?: FileStamp) {
    this.handle = handle;
    this.options = {
      pollMs: 1500,
      settleDelayMs: 400,
      ...options,
    };
    this.lastStamp = initialStamp ?? null;
  }

  /** Cập nhật dấu vân tay hiện tại — gọi sau khi app tự đọc/tự ghi file (để không tự báo "đổi" do chính mình gây ra) */
  setStamp(stamp: FileStamp): void {
    this.lastStamp = stamp;
  }

  start(): void {
    if (!this.stopped) return;
    this.stopped = false;
    this.timerId = (globalThis as any).setInterval(() => this.tick(), this.options.pollMs);
  }

  stop(): void {
    this.stopped = true;
    if (this.timerId != null) {
      (globalThis as any).clearInterval(this.timerId);
      this.timerId = null;
    }
  }

  private async tick(): Promise<void> {
    if (this.stopped || this.busy) return;
    this.busy = true;
    try {
      const file = await this.handle.getFile();
      const stamp = stampOf(file);
      if (!stampChanged(this.lastStamp, stamp)) return; // không đổi -> không làm gì thêm

      // Chờ file "ổn định" — Excel/Office có thể ghi file qua nhiều bước
      await new Promise((r) => (globalThis as any).setTimeout(r, this.options.settleDelayMs));
      if (this.stopped) return;

      const fileAgain = await this.handle.getFile();
      const stampAgain = stampOf(fileAgain);
      if (stampAgain.mtime !== stamp.mtime || stampAgain.size !== stamp.size) {
        return; // vẫn đang ghi dở -> để nhịp poll sau xử lý tiếp
      }

      if (this.options.shouldDefer?.()) {
        this.options.onPendingChange?.();
        return; // có thay đổi chưa lưu trên app -> không tự ghi đè, chỉ báo hiệu
      }

      this.lastStamp = stampAgain;
      await this.options.onChanged({ file: fileAgain, stamp: stampAgain });
    } catch (err) {
      // Quyền bị thu hồi / file bị khoá / đang ghi -> bỏ qua nhịp này, thử lại nhịp sau
      this.options.onError?.(err);
    } finally {
      this.busy = false;
    }
  }
}

// ----------------------------------------------------------------------------
// 5. (TUỲ CHỌN) REACT HOOK bọc quanh FileHandleWatcher — xoá phần này nếu
//    project khác không dùng React, phần lõi ở trên vẫn dùng độc lập được.
// ----------------------------------------------------------------------------
/*
import { useEffect, useRef, useState } from 'react';

export function useFileHandleWatcher(
  handle: FileSystemFileHandle | null,
  enabled: boolean,
  onChanged: (file: File, buffer: ArrayBuffer) => void,
  deps: { hasUnsavedChanges: boolean }
) {
  const [pendingChange, setPendingChange] = useState(false);
  const [lastSyncAt, setLastSyncAt] = useState<Date | null>(null);
  const watcherRef = useRef<FileHandleWatcher | null>(null);

  useEffect(() => {
    if (!handle || !enabled) return;

    const watcher = new FileHandleWatcher(handle, {
      pollMs: 1500,
      settleDelayMs: 400,
      shouldDefer: () => deps.hasUnsavedChanges,
      onPendingChange: () => setPendingChange(true),
      onChanged: async ({ file }) => {
        const buffer = await file.arrayBuffer();
        onChanged(file, buffer);
        setPendingChange(false);
        setLastSyncAt(new Date());
      },
      onError: (err) => console.warn('[FileHandleWatcher]', err),
    });
    watcherRef.current = watcher;
    watcher.start();

    return () => watcher.stop();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [handle, enabled, deps.hasUnsavedChanges]);

  return { pendingChange, lastSyncAt };
}
*/

// ----------------------------------------------------------------------------
// 6. VÍ DỤ SỬ DỤNG THUẦN (vanilla, không React)
// ----------------------------------------------------------------------------
/*
const EXCEL_TYPES: FilePickerAcceptType[] = [
  {
    description: 'Excel workbook',
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
    },
  },
];

async function main() {
  const opened = await openFileWithHandle(EXCEL_TYPES, 64);
  if (!opened) return; // user huỷ hoặc trình duyệt không hỗ trợ

  let hasUnsavedUiChanges = false; // cờ này do app tự quản lý

  const watcher = new FileHandleWatcher(
    opened.handle,
    {
      pollMs: 1500,
      settleDelayMs: 400,
      shouldDefer: () => hasUnsavedUiChanges,
      onPendingChange: () => console.log('File trên đĩa đã đổi nhưng đang có sửa chưa lưu trên app.'),
      onChanged: async ({ file }) => {
        const buffer = await file.arrayBuffer();
        console.log('File vừa được lưu lại trên đĩa, đã đọc lại:', file.name, buffer.byteLength, 'bytes');
        // TODO: parse buffer bằng logic của project khác, cập nhật UI...
      },
      onError: (err) => console.warn('Poll lỗi (bỏ qua nhịp này):', err),
    },
    stampOf(opened.file)
  );

  watcher.start();
  // watcher.stop(); // gọi khi unmount / đóng file
}
*/
