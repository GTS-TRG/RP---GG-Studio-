import { CableSpecRow } from '../types';

/**
 * Default Cable Standard Specification Table (Spec. Cable)
 * Standard TCVN / IEC cable cross-sections for circuit breakers
 */
export const DEFAULT_CABLE_SPECS: CableSpecRow[] = [
  { cbAmp: 6, phaseText: '1.5', peText: '1.5' },
  { cbAmp: 10, phaseText: '1.5', peText: '1.5' },
  { cbAmp: 16, phaseText: '2.5', peText: '2.5' },
  { cbAmp: 20, phaseText: '4', peText: '4' },
  { cbAmp: 25, phaseText: '4', peText: '4' },
  { cbAmp: 32, phaseText: '6', peText: '6' },
  { cbAmp: 40, phaseText: '10', peText: '10' },
  { cbAmp: 50, phaseText: '10', peText: '10' },
  { cbAmp: 63, phaseText: '16', peText: '16' },
  { cbAmp: 80, phaseText: '25', peText: '16' },
  { cbAmp: 100, phaseText: '35', peText: '16' },
  { cbAmp: 125, phaseText: '50', peText: '25' },
  { cbAmp: 160, phaseText: '70', peText: '35' },
  { cbAmp: 200, phaseText: '95', peText: '50' },
  { cbAmp: 250, phaseText: '120', peText: '70' },
  { cbAmp: 315, phaseText: '185', peText: '95' },
  { cbAmp: 400, phaseText: '240', peText: '120' },
  { cbAmp: 500, phaseText: '2x150', peText: '150' },
  { cbAmp: 630, phaseText: '2x185', peText: '185' },
  { cbAmp: 800, phaseText: '2x240', peText: '240' },
  { cbAmp: 1000, phaseText: '3x240', peText: '240' },
  { cbAmp: 1250, phaseText: '4x240', peText: '240' },
  { cbAmp: 1600, phaseText: '5x240', peText: '240' },
];

/**
 * Finds required CB rating from spec table for a given calculated load in Amperes
 */
export function getRequiredCBFromSpec(reqAmp: number, specs: CableSpecRow[]): number {
  for (const spec of specs) {
    if (reqAmp <= spec.cbAmp) {
      return spec.cbAmp;
    }
  }
  return specs.length > 0 ? specs[specs.length - 1].cbAmp : reqAmp;
}

/**
 * Gets index (0-based) of the spec row matching or exceeding the given CB Ampere
 */
export function getSpecIndex(cbAmp: number, specs: CableSpecRow[]): number {
  for (let i = 0; i < specs.length; i++) {
    if (cbAmp <= specs[i].cbAmp) {
      return i;
    }
  }
  return specs.length > 0 ? specs.length - 1 : -1;
}
