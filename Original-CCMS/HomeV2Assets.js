/**
 * CCMS HOME V2 — APPROVED VISUAL ASSET ADAPTER
 * Candidate shell support only. No production-data writes. No doGet().
 */
const CCMS_HOME_V2_ASSETS = Object.freeze({
  BRAND_XRAY:            '1cGSUZmN0z_LJR9JFbphwgD24i6JzVjFr',
  HOME_HERO:             '1fQE23L4375u46Mrmph_WxFTcsc-A13h-',
  HOME_FLAGGED:          '1BHPWsyzSMk13p38urUb6OjY3fdcSS-tq',
  HOME_WAITING_REVIEW:   '13A2UwvVQeN92WY2_VArV-HcqzcN59eKD',
  HOME_REVIEWED:         '1IMDJZhwlnFniD3x8Kfubo6pfwtHy4Osv',
  HOME_TODAYS_SCHEDULE:  '1CVx_ovVaJ-fZLhUJ1pcSG-uHld_27Wqs',
  QUICK_STUDENTS:        '1GaJyYharOhMVKlDoHne1qROagGEFlHLW',
  QUICK_INSTRUCTORS:     '12R58JINeYXZhMJZHRSWvczZO7xDZXH0v',
  QUICK_SITES:           '15_hEdE-DaDPGQ4KXWS3JH2nN-bEWOGAL',
  QUICK_SCHEDULING:      '1ieHk3SlCMydpcyPbXT2TXe1B0YondNe4',
  QUICK_PAPERWORK:       '1-ysD1PKptxNIr2PZpHdlnrrnTaA-043E',
  QUICK_SYSTEM_SETUP:    '18f-Tg3sc_yQz2vuKL_7ZZLwBqPm4FcA1'
});

function dashboardHomeV2GetAssets() {
  const output = {};
  Object.keys(CCMS_HOME_V2_ASSETS).forEach(function (key) {
    const blob = DriveApp.getFileById(CCMS_HOME_V2_ASSETS[key]).getBlob();
    const mimeType = blob.getContentType() || 'image/jpeg';
    output[key] = 'data:' + mimeType + ';base64,' +
      Utilities.base64Encode(blob.getBytes());
  });
  return output;
}

function testDashboardHomeV2Assets() {
  const result = { success: true, assets: [], failures: [] };
  Object.keys(CCMS_HOME_V2_ASSETS).forEach(function (key) {
    try {
      const file = DriveApp.getFileById(CCMS_HOME_V2_ASSETS[key]);
      result.assets.push({
        key: key, fileId: file.getId(), name: file.getName(),
        mimeType: file.getMimeType(), size: file.getSize()
      });
    } catch (err) {
      result.success = false;
      result.failures.push({
        key: key, fileId: CCMS_HOME_V2_ASSETS[key],
        error: String(err && err.message ? err.message : err)
      });
    }
  });
  return result;
}
