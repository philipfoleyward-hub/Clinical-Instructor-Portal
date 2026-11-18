/** Candidate image service — standalone CI Portal. No CCMS library dependency. */
const CI_PORTAL_IMAGE_ASSETS = Object.freeze({
  LOGIN_BACKGROUND: '1WQXvR1QCDmMhHs60HzmLGmAjNkO3VJAw',
  HOME_BANNER: '1JnfVDEtaQunPzKla2MTaM6MYTdwqk3rp',
  // Existing candidate logo: verify against approved branding before production.
  PORTAL_LOGO: '122DmjnUCaxn-aKB3oj87XHPqyMXsbeFY'
});

/** Same server-to-client data-URL pattern as DashboardrunV2. */
function ciPortalGetAssets() {
  const result = {};
  Object.keys(CI_PORTAL_IMAGE_ASSETS).forEach(function(key) {
    const blob = DriveApp.getFileById(CI_PORTAL_IMAGE_ASSETS[key]).getBlob();
    const mime = blob.getContentType() || 'image/png';
    if (mime.indexOf('image/') !== 0) throw new Error('Not an image: ' + key);
    result[key] = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
  });
  return result;
}

/** Read-only diagnostic; run manually to verify permissions and file metadata. */
function testCiPortalImageAssets() {
  const result = {success:true, assets:[], failures:[]};
  Object.keys(CI_PORTAL_IMAGE_ASSETS).forEach(function(key) {
    try {
      const file = DriveApp.getFileById(CI_PORTAL_IMAGE_ASSETS[key]);
      result.assets.push({key:key, fileId:file.getId(), name:file.getName(),
        mimeType:file.getMimeType(), size:file.getSize()});
    } catch(err) {
      result.success=false;
      result.failures.push({key:key, error:String(err&&err.message||err)});
    }
  });
  return result;
}
