import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

// 対象ファイルパス
const filePath = path.join(process.cwd(), 'samples', '【D外設】【24-01】ＦＡＱ.xls');

console.log(`Analyzing file: ${filePath}`);

try {
  // 1. ファイル読み込み
  const buffer = fs.readFileSync(filePath);
  const workbook = XLSX.read(buffer, { type: 'buffer' });

  console.log('Successfully read workbook.');
  console.log(`Sheet Names: ${workbook.SheetNames.join(', ')}`);

  // 2. 最初のシートを解析
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  if (!worksheet) {
    console.error('Error: First sheet is empty or invalid.');
    process.exit(1);
  }

  console.log(`Analyzing Sheet: "${firstSheetName}"`);

  // 3. 範囲情報の確認
  const range = worksheet['!ref'];
  console.log(`Range: ${range}`);

  // 4. HTML変換テスト
  const html = XLSX.utils.sheet_to_html(worksheet);
  
  // 5. 問題点の洗い出し分析
  console.log('\n--- Analysis Report ---');

  // a. 結合セルの有無
  const merges = worksheet['!merges'] || [];
  console.log(`Merged Cells: ${merges.length} found.`);
  if (merges.length > 0) {
    console.log('  -> Warning: Complex layout with merged cells may cause rendering issues in HTML.');
  }

  // b. 画像・図形の検出 (SheetJS Free版では画像は通常取得できないが、Drawingオブジェクトの痕跡を確認)
  // imagesプロパティなどはPro版機能だが、キーとして何らかのオブジェクトがあるか確認
  const hasImages = Object.keys(worksheet).some(key => key.startsWith('!drawing') || key.startsWith('!images'));
  if (hasImages) {
     console.log('  -> Warning: Drawing/Image objects detected. These will likely NOT be rendered in the output image.');
  } else {
     console.log('  -> Note: No standard drawing objects detected (or not accessible via Community Edition).');
  }

  // c. HTML構造チェック
  const tableMatch = html.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
  if (tableMatch) {
    console.log('  -> Table HTML generated successfully.');
    // 行数の概算
    const rowCount = (html.match(/<tr/g) || []).length;
    console.log(`  -> Estimated Rows: ${rowCount}`);
    
    if (rowCount > 100) {
      console.log('  -> Warning: Large number of rows. Generating a single image may result in a very tall image or memory issues.');
    }
  } else {
    console.error('  -> Error: Failed to generate HTML table structure.');
  }

  // d. データ内容の簡易チェック (文字化け等の確認のため、一部を出力)
  // セルの値をいくつかサンプリング
  console.log('\n--- Content Sample (First 5 non-empty cells) ---');
  let count = 0;
  for (let key in worksheet) {
    if (key.startsWith('!') || count >= 5) continue;
    console.log(`  [${key}]: ${JSON.stringify(worksheet[key].v)} (formatted: ${worksheet[key].w})`);
    count++;
  }

} catch (err) {
  console.error('Analysis Failed:', err);
}
