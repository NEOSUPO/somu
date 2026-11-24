function autoBackup() {
  // 1. バックアップ元のスプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileId = ss.getId();
  const file = DriveApp.getFileById(fileId);

  // 2. バックアップ先のフォルダIDを指定（※ここにメモしたフォルダIDを貼り付けてください）
  const folderId = '1OsG4dzRXwN4goEqLNF2fWWAcQ0lVcd66';
  const folder = DriveApp.getFolderById(folderId);

  // 3. バックアップ時のファイル名を設定（例：元のファイル名_20260508_010000）
  const date = new Date();
  const dateString = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = ss.getName() + '_' + dateString;

  // 4. 指定したフォルダにコピーを作成
  file.makeCopy(fileName, folder);
}