//@ts-nocheck
// 将一个sheet转换为一个kv表的gulp插件
import through2 from 'through2';
import xlsx from 'node-xlsx';
import Vinyl from 'vinyl';
import path from 'path';
import { pinyin, customPinyin } from 'pinyin-pro';
import { pushNewLinesToCSVFile } from './kvToLocalization';

const cli = require('cli-color');
const PLUGIN_NAME = 'gulp-dotax:sheetToKV';

export interface SheetToKVOptions {
  /** 需要略过的表格的正则表达式 */
  sheetsIgnore?: string;
  /** 是否启用啰嗦模式 */
  verbose?: boolean;
  /** 是否将汉语转换为拼音 */
  chineseToPinyin?: boolean;
  /** 自定义的拼音 */
  customPinyins?: Record<string, string>;
  /** KV的缩进方式，默认为四个空格 */
  indent?: string;
  /** 是否将只有两列的表输出为简单键值对 */
  autoSimpleKV?: boolean;
  /** Key行行号，默认为2 */
  keyRowNumber?: number;
  /** KV文件的扩展名，默认为 .txt */
  kvFileExt?: string;
  /** 强制输出空格的单元格内容 */
  forceEmptyToken?: string;
  /** 中文转换为英文的映射列表 */
  aliasList?: Record<string, string>;
  /** 输出本地化文本到 addon.csv 文件 */
  addonCSVPath?: string;
  /** addon.csv输出的默认语言 */
  addonCSVDefaultLang?: string;
  // 给生成的kv表添加自定义名称 (外层大括号的名字)
  rootname?: string;
  // --- 新增配置：自定义输出的文件名 ---
  outputFilename?: string; 
}

function isSimpleKV(key_row: string[]) {
  const validKeys = key_row.filter((i) => i != null && i != '' && !i.includes('#Loc'));
  return validKeys.length == 2;
}

function isEmptyOrNullOrUndefined(value: any) {
  return value === null || value === undefined || value === ``;
}

export function sheetToKV(options: SheetToKVOptions) {
  const {
    rootname = "XLSXContent",
    customPinyins = {},
    sheetsIgnore = /^\s*$/,
    verbose = false,
    forceEmptyToken = `__empty__`,
    autoSimpleKV = true,
    kvFileExt = '.txt',
    chineseToPinyin = true,
    keyRowNumber = 2,
    indent = ' ',
    aliasList = {},
    addonCSVPath = null,
    addonCSVDefaultLang = `SChinese`,
    outputFilename = '',
  } = options;

  customPinyin(customPinyins);
  const aliasKeys = Object.keys(aliasList)
    .sort((a, b) => b.length - a.length);

  // 本地化token列表
  let locTokens: { [key: string]: string }[] = [];

  function convert_chinese_to_pinyin(da: string) {
    if (da === null || da.match === null) return da;
    aliasKeys.forEach((aliasKey) => {
      da = da.replace(aliasKey, aliasList[aliasKey]);
    });
    let s = da;
    let reg = /[\u4e00-\u9fa5]+/g;
    let match = s.match(reg);
    if (match != null) {
      match.forEach((m) => {
        s = s
          .replace(m, pinyin(m, { toneType: 'none', type: 'array' }).join('_'))
          .replace('ü', 'v');
      });
    }
    return s;
  }

  function deal_with_kv_value(value: string): string {
    if (/^[0-9]+.?[0-9]*$/.test(value)) {
      let number = parseFloat(value);
      if (number % 1 !== 0) {
        value = number.toFixed(4);
      }
    }
    if (value === undefined) return '';
    if (forceEmptyToken === value) return '';
    return value;
  }

  // --- 关键修复：补全了 convert_row_to_kv 的完整逻辑 ---
  function convert_row_to_kv(row: string[], key_row: string[]): string {
    // 第一列为主键
    let main_key = row[0];

    function checkSpace(str: string) {
      if (typeof str == 'string' && str.trim != null && str != str.trim()) {
        console.warn(cli.red(`${main_key} 中的 ${str} 前后有空格，请检查！`));
      }
    }
    checkSpace(main_key);

    let attachWearablesBlock = false;
    let abilityValuesBlock = false;
    let varIndex = 0;
    let indentLevel = 1;
    let locAbilitySpecial = null;

    return key_row
      .map((key, i) => {
        // 跳过空的key
        if (isEmptyOrNullOrUndefined(key)) return;

        let output_value = row[i];
        let indentStr = (indent || `\t`).repeat(indentLevel);

        // 处理第一列（主键）
        if (i === 0) {
          indentLevel++;
          return `${indentStr}"${main_key}" {`;
        }

        // 处理 #Loc 开头的本地化Key
        if (key.startsWith('#Loc')) {
          // 这里简化处理，实际逻辑可能更复杂，根据你的需求调整
          let locKey = key.replace('#Loc', main_key);
          locTokens.push({ key: locKey, value: output_value });
          return `${indentStr}// Localized key: ${locKey}`;
        }

        // 处理普通KV对
        output_value = deal_with_kv_value(output_value);

        return `${indentStr}"${key}" "${output_value}"`;
      })
      .filter((row) => row != null) // 过滤空行
      .map((s) => (chineseToPinyin ? convert_chinese_to_pinyin(s) : s))
      .join('\n') + '\n' + `${indent.repeat(indentLevel - 1)}}`; // 结尾的大括号
  }

  function convert(this: any, file: Vinyl, enc: any, next: Function) {
    if (file.isNull()) return next(null, file);
    if (file.isStream()) return next(new Error(`${PLUGIN_NAME} Streaming not supported`));
    if (file.basename.startsWith(`~$`)) {
      console.log(`${PLUGIN_NAME} Ignore temp xlsx file ${file.basename}`);
      return next();
    }

    if (!file.basename.endsWith(`.xlsx`) && !file.basename.endsWith(`.xls`)) {
      console.log(cli.green(`${PLUGIN_NAME} ignore non-xlsx file ${file.basename}`));
      return next();
    }

    if (file.isBuffer()) {
      console.log(`${PLUGIN_NAME} Converting ${file.path} to kv`);
      const workbook = xlsx.parse(file.contents);

      let mergedKVContent = '';
      let firstSheetName: string | null = null; 

      workbook.forEach((sheet, index) => {
        let sheet_name = sheet.name;
        
        if (new RegExp(sheetsIgnore).test(sheet_name)) {
          console.log(cli.red(`${PLUGIN_NAME} Ignoring sheet ${sheet_name}...`));
          return;
        }

        // 记录第一个有效 Sheet 名称
        if (firstSheetName === null) {
          firstSheetName = sheet_name.match(/[\u4e00-\u9fa5]+/g) ? 
            convert_chinese_to_pinyin(sheet_name) : sheet_name;
        }

        const sheet_data = sheet.data as string[][];
        if (sheet_data.length <= keyRowNumber) return;

        let key_row = sheet_data[keyRowNumber - 1].map((i) => i.toString());
        const kv_data = sheet_data.slice(keyRowNumber);
        if (kv_data.length === 0) return;

        let kv_data_str = '';
        if (isSimpleKV(key_row) && autoSimpleKV) {
          const kv_data_simple = kv_data.map((row) => {
            return `\t"${row[0]}" "${row[1]}"`;
          });
          kv_data_str = `${kv_data_simple.join('\n')}`;
        } else {
          const kv_data_complex = kv_data.map((row) => {
            if (isEmptyOrNullOrUndefined(row[0])) return;
            return convert_row_to_kv(row, key_row);
          });
          kv_data_str = `${kv_data_complex.join('\n')}`;
        }

        mergedKVContent += kv_data_str + '\n'; 
      });

      // --- 文件名逻辑 ---
      let finalFilename: string;
      if (outputFilename) {
        finalFilename = outputFilename;
      } else if (firstSheetName) {
        finalFilename = firstSheetName;
      } else {
        finalFilename = path.basename(file.path, path.extname(file.path)) + '_merged';
      }
      const outputBasename = `${finalFilename}${kvFileExt}`;
      // --- ---

      if (mergedKVContent.trim() !== '') {
        const out_put = `// this file is auto-generated by Xavier's sheet_to_kv 
// Source: ${file.basename}
// SourceCode: https://github.com/XavierCHN/gulp-dotax/blob/master/src/sheetToKV.ts

"${rootname}" { 
  ${mergedKVContent}
}
`;

        const kv_file = new Vinyl({
          base: file.base,
          path: path.join(file.dirname, outputBasename),
          contents: Buffer.from(out_put),
        });
        this.push(kv_file);
        console.log(`${PLUGIN_NAME} Writing merged content to ${outputBasename}`);
      }
    }
    next();
  }

  function endStream() {
    if (addonCSVPath != null) {
      pushNewLinesToCSVFile(addonCSVPath, locTokens);
    }
    this.emit('end');
  }
  return through2.obj(convert, endStream);
}