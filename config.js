// Steak Teppei Payroll — shared config (v1.0)
// Single source of truth for the GAS endpoint + access secret.
//
// After every new GAS deployment, change GAS_URL HERE ONLY.
// Both payroll.html and payroll-summary.html read from this file,
// so you no longer edit the URL in two places.
//
// GAS_SECRET must match the SECRET constant inside gas_payroll.js.
// Requests without the correct secret are rejected by the GAS side.
window.ST_CONFIG = {
  GAS_URL:    'https://script.google.com/macros/s/AKfycbxPOnlUvOjp9oFCz96vexbYE324mDKULtz2fUbdmJQTFUuB-M8rRgGapA5_ICqKO8cD/exec',
  GAS_SECRET: 'st_fbd09725f529893def93cfee7cbb203d'
};
