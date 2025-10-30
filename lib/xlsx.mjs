// ES Module wrapper for xlsx.bundle.js (UMD)
// This file imports the UMD bundle and re-exports XLSX for ES module usage

// Since the UMD bundle sets window.XLSX, we need to ensure it's loaded
// We'll use dynamic import to load the script, then export the global XLSX

// Import the UMD bundle (it will set window.XLSX)
import './xlsx.bundle.js';

// Wait a tick for the global to be set, then export it
const XLSX = window.XLSX;

if (!XLSX) {
    throw new Error('XLSX failed to load from xlsx.bundle.js');
}

export default XLSX;
export { XLSX };
