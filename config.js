const CONFIG = {

    /*
     * Use the SAME Google Web Client ID as your current working FM dashboard.
     */
    GOOGLE_CLIENT_ID: "50542963972-75ictdqcblqigj1iq30pobdqk06s1v93.apps.googleusercontent.com",

    /*
     * The Assets Inventory Ledger is the only operational workbook used by
     * the dashboard. Routine movements and persistent alerts live here too.
     */
    ROUTINE_SHEET_NAME:
        "Routine Outbounds",

    ALERTS_SHEET_NAME:
        "Alerts",

    /*
     * Assets Inventory Ledger - dedicated spreadsheet (opened directly by
     * ID, not searched by name) that is the single read/write source for
     * inventory and asset-movement data. INVENTORY_SHEET_NAME and
     * TRANSACTIONS_SHEET_NAME below both live in THIS spreadsheet.
     */
    INVENTORY_LEDGER_SHEET_ID:
        "1VC44seK6vR2IHl53Bn0NwentqRd8XwhSZJ0RtWB67uU",

    INVENTORY_LEDGER_NAME:
        "Assets Inventory Ledger",

    INVENTORY_SHEET_NAME:
        "Asset Inventory",

    TRANSACTIONS_SHEET_NAME:
        "Asset Transactions",

    CLIENT_LIST_SHEET_NAME:
        "Client List",

    LOCATION:
        "UB11",

    /*
     * The "self" client used for DISCARD movements (assets scrapped/thrown
     * away rather than sent to or received from an outside client). Must
     * match dashboard.js's discard handling exactly so movements recorded
     * from either page show up as the same client in the audit trail.
     */
    SELF_CLIENT_NAME:
        "HSC London (Self)",

    /*
     * Parent Google Drive folder that movement photos are filed under.
     * A new sub-folder named with the day's date (YYYY-MM-DD) is created
     * inside this folder automatically the first time a photo is uploaded
     * on a given day, and every photo taken that day is stored there.
     * https://drive.google.com/drive/folders/15NJveQRQ0WHR5FuR0E967dGs5bOBaPoH
     */
    DRIVE_PHOTOS_PARENT_FOLDER_ID:
        "15NJveQRQ0WHR5FuR0E967dGs5bOBaPoH",

    /*
     * NOTE: broadened from "drive.readonly" to full "drive" so the app can
     * create the daily date-folders and upload photos into them. Existing
     * signed-in users will be prompted once to re-grant access because the
     * scope changed.
     */
    OAUTH_SCOPES: [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ].join(" ")

};

/* Compatibility helper.
   Some older pages call formatTimestamp() outside their own module wrapper.
   Defining it here prevents "formatTimestamp is not defined" when config.js
   is loaded before those pages. */
window.formatTimestamp = window.formatTimestamp || function(value){
    if(!value) return "";
    const d = new Date(value);
    return Number.isNaN(d.getTime())
        ? String(value)
        : d.toLocaleString("en-GB", {
            day:"2-digit",
            month:"short",
            hour:"2-digit",
            minute:"2-digit"
        });
};


/* Shared Google auth cache. Keeps a short-lived access token across page changes/refreshes
   so the user normally does not have to reconnect Sheets on every page. */
window.FM_AUTH_CACHE = window.FM_AUTH_CACHE || (() => {
    const KEY = "fmAssetAccessToken";
    const SKEW_MS = 45 * 1000;
    function read(expectedEmail){
        try{
            const value = JSON.parse(localStorage.getItem(KEY) || "null");
            if(expectedEmail && value?.email && String(value.email).toLowerCase() !== String(expectedEmail).toLowerCase()){
                localStorage.removeItem(KEY);
                return null;
            }
            if(!value?.token || Number(value.expiresAt || 0) <= Date.now() + SKEW_MS){
                localStorage.removeItem(KEY);
                return null;
            }
            return value;
        }catch(_){
            localStorage.removeItem(KEY);
            return null;
        }
    }
    function write(token, expiresIn, email){
        try{
            localStorage.setItem(KEY, JSON.stringify({
                token,
                expiresAt: Date.now() + Math.max(60, Number(expiresIn || 3600)) * 1000,
                email: email || ""
            }));
        }catch(_){}
    }
    function clear(){ try{ localStorage.removeItem(KEY); }catch(_){} }
    return {read,write,clear};
})();

/* Shared media helpers used by the dashboard, movement logger and alerts. */
window.FM_MEDIA = window.FM_MEDIA || (() => {
    const DRIVE_API = "https://www.googleapis.com/drive/v3/files";

    function driveFileIdFromLink(url){
        const text = String(url || "");
        const m = text.match(/\/d\/([a-zA-Z0-9_-]+)/) || text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
        return m ? m[1] : "";
    }

    function revokeObjectUrl(img){
        const previous = img?.dataset?.fmObjectUrl;
        if(previous){
            try { URL.revokeObjectURL(previous); } catch (_) {}
            delete img.dataset.fmObjectUrl;
        }
    }

    async function loadDriveImage(img, url, accessToken){
        if(!img) return;
        revokeObjectUrl(img);
        if(!url){
            img.removeAttribute("src");
            img.dataset.fmImageState = "empty";
            return false;
        }
        const fileId = driveFileIdFromLink(url);
        if(!fileId || !accessToken){
            img.src = url;
            img.dataset.fmImageState = "direct";
            return true;
        }

        img.dataset.fmImageState = "loading";
        try{
            const response = await fetch(`${DRIVE_API}/${encodeURIComponent(fileId)}?alt=media`, {
                headers: { Authorization: `Bearer ${accessToken}` },
                cache: "force-cache"
            });
            if(!response.ok) throw new Error(`Drive image request failed (${response.status})`);
            const blob = await response.blob();
            const objectUrl = URL.createObjectURL(blob);
            img.dataset.fmObjectUrl = objectUrl;
            img.src = objectUrl;
            img.dataset.fmImageState = "loaded";
            return true;
        }catch(error){
            console.warn("Authenticated Drive image load failed:", error);
            try{
                const fallback = await fetch(`https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=w1200`, {
                    headers: { Authorization: `Bearer ${accessToken}` },
                    cache: "force-cache"
                });
                if(fallback.ok){
                    const blob = await fallback.blob();
                    const objectUrl = URL.createObjectURL(blob);
                    img.dataset.fmObjectUrl = objectUrl;
                    img.src = objectUrl;
                    img.dataset.fmImageState = "thumbnail";
                    return true;
                }
            }catch(_){ /* fall through to Drive link */ }
            img.src = url;
            img.dataset.fmImageState = "fallback";
            return false;
        }
    }

    async function optimizeImageForUpload(file, { maxDimension = 1800, quality = 0.82 } = {}){
        if(!file || !String(file.type || "").startsWith("image/")) return file;
        try{
            const objectUrl = URL.createObjectURL(file);
            try{
                const image = new Image();
                image.decoding = "async";
                image.src = objectUrl;
                await new Promise((resolve, reject) => {
                    image.onload = resolve;
                    image.onerror = reject;
                });
                const scale = Math.min(1, maxDimension / Math.max(image.naturalWidth || image.width, image.naturalHeight || image.height));
                const width = Math.max(1, Math.round((image.naturalWidth || image.width) * scale));
                const height = Math.max(1, Math.round((image.naturalHeight || image.height) * scale));
                const canvas = document.createElement("canvas");
                canvas.width = width;
                canvas.height = height;
                const ctx = canvas.getContext("2d", { alpha: false });
                if(!ctx) return file;
                ctx.drawImage(image, 0, 0, width, height);
                const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/jpeg", quality));
                if(!blob || blob.size >= file.size * 0.92) return file;
                const base = String(file.name || "movement").replace(/\.[^.]+$/, "");
                return new File([blob], `${base}.jpg`, { type: "image/jpeg", lastModified: Date.now() });
            } finally {
                URL.revokeObjectURL(objectUrl);
            }
        }catch(error){
            console.warn("Image optimisation skipped:", error);
            return file;
        }
    }

    return { driveFileIdFromLink, loadDriveImage, optimizeImageForUpload, revokeObjectUrl };
})();
