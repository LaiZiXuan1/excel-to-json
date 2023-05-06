import xlsx from 'node-xlsx';
import fs from 'fs-extra';
import path from 'path';

const root = path.resolve(__dirname, '../', 'index');
const outDir = path.resolve(__dirname, '../', 'data');
const db: any = {};

start();

interface Sheet {
  name: string;
  data: unknown[];
}

// 读取 xlsx
async function readXlsx(filepath: any) {
  console.log(`【${filepath}】 开始转换`);

  const filename = path.basename(filepath, path.extname(filepath));

  const sheets = xlsx.parse(filepath);

  if (sheets.length === 0) return;

  if (sheets.length === 1) {
    db[filename] = handleSheet(sheets[0]);
  } else {
    db[filename] = {};
    for await (const sheet of sheets) {
      db[filename][sheet.name] = handleSheet(sheet);
    }
  }

  function handleSheet(sheet: Sheet) {
    const { data } = sheet;

    const result = [] as Array<Record<string, any>>;

    if (data.length <= 1) return result;

    const keys = data[0] as Array<string>;

    data.slice(1, data.length).forEach((line: any[]) => {
      const obj: Record<string, any> = {};

      keys.forEach((key, index) => {
        obj[key] = line[index];
      });

      result.push(obj);
    });

    return result;
  }

  console.log(`【${filepath}】 转换完毕`);
}

// 读取目录
async function readdir(dirname: any) {
  const direntArr = fs.readdirSync(dirname, { withFileTypes: true });
  for await (const dirent of direntArr) {
    if (dirent.name.startsWith('~$')) continue;
    if (dirent.name.startsWith('ignore#')) continue;

    const _path = path.resolve(dirname, dirent.name);

    if (dirent.isDirectory()) {
      await readdir(_path);
    }

    if (dirent.isFile()) {
      if (['.xlsx', '.xls', '.csv'].includes(path.extname(_path))) {
        await readXlsx(_path);
      }
    }
  }
}

async function start() {
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
  await readdir(root);
  fs.writeFileSync(path.resolve(outDir, `data.json`), JSON.stringify(db));

  console.log('完成了！！！');
}
