/**
 * スクリプトプロパティから指定されたキーの値を取得する共通関数。
 * 値が設定されていない場合はエラーをログに記録し、nullを返す。
 * * @param {string} key 取得したいプロパティのキー (例: 'ADMIN_CALENDAR_ID')
 * @returns {string | null} 取得した値、または設定されていない場合は null
 */
function getRequiredProperty(key) {
    const value = SCRIPT_PROPERTIES.getProperty(key);

    if (!value || value.trim() === '') {
        Logger.log(`[ERROR] 必須のスクリプトプロパティが見つかりません: ${key}`);
        // 開発環境で動作を止める場合は throw new Error に変更しても良い
        return null;
    }
    return value.trim();
}