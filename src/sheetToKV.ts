//@ts-nocheck
import through2 from 'through2';
import xlsx from 'node-xlsx';
import Vinyl from 'vinyl';
import path from 'path';
import { pinyin, customPinyin } from 'pinyin-pro';
import { pushNewLinesToCSVFile } from './kvToLocalization';

const cli = require('cli-color');
const PLUGIN_NAME = 'gulp-dotax:sheetToKV';

export interface SheetToKVOptions {
  sheetsIgnore?: string;
  verbose?: boolean;
  chineseToPinyin?: boolean;
  customPinyins?: Record<string, string>;
  indent?: string;
  autoSimpleKV?: boolean;
  keyRowNumber?: number;
  kvFileExt?: string;
  forceEmptyToken?: string;
  aliasList?: Record<string, string>;
  addonCSVPath?: string;
  addonCSVDefaultLang?: string;
  rootname?: string;
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
    indent = '    ', // 默认4个空格
    aliasList = {},
    addonCSVPath = null,
    addonCSVDefaultLang = `SChinese`,
    outputFilename = '',
  } = options;

  customPinyin(customPinyins);
  const aliasKeys = Object.keys(aliasList).sort((a, b) => b.length - a.length);
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
        s = s.replace(m, pinyin(m, { toneType: 'none', type: 'array' }).join('_')).replace('ü', 'v');
      });
    }
    return s;
  }

  function deal_with_kv_value(value: string): string {
    if (/^[0-9]+.?[0-9]*$/.test(value)) {
      let number = parseFloat(value);
      if (number % 1 !== 0) value = number.toFixed(4);
    }
    if (value === undefined) return '';
    if (forceEmptyToken === value) return '';
    return value;
  }

  // --- 彻底重写的 convert_row_to_kv ---
  // 逻辑：先收集内容，最后统一组装，确保大括号层级正确
  function convert_row_to_kv(row: string[], key_row: string[]): string {
    let main_key = row[0];
    
    // 检查空格
    if (typeof main_key == 'string' && main_key.trim() !== main_key) {
        console.warn(cli.red(`${main_key} 前后有空格，请检查！`));
    }

    let attachWearablesBlock = false;
    let abilityValuesBlock = false;
    let varIndex = 0;
    let locAbilitySpecial = null;
    
    const baseIndent = indent || '\t';
    const innerIndent = baseIndent + baseIndent; // 内部缩进

    // 1. 收集该行的所有子键值对
    let contentLines: string[] = [];

    key_row.forEach((key, i) => {
      if (isEmptyOrNullOrUndefined(key)) return;
      
      // 跳过第一列（主键），只处理数据列
      if (i === 0) return; 

      let output_value = row[i];

      // --- 状态机逻辑 ---
      if (key === `AttachWearables[{]`) attachWearablesBlock = true;
      if (attachWearablesBlock && key == `}]`) attachWearablesBlock = false;
      if (key === `AbilityValues[{]`) abilityValuesBlock = true;
      if (abilityValuesBlock && key === `}]`) abilityValuesBlock = false;

      // --- 处理特殊块内容 ---
      
      // 1. 饰品
      if (attachWearablesBlock && key !== `AttachWearables[{]`) {
         if (output_value != `` && output_value != undefined) {
           if (output_value.toString().trimStart().startsWith('{')) {
             contentLines.push(`${innerIndent}"${key}" ${output_value}`);
           } else {
             contentLines.push(`${innerIndent}"${key}" { "ItemDef" "${output_value}" }`);
           }
         }
         return;
      }

      // 2. 技能数值
      if (abilityValuesBlock && key !== `AbilityValues[{]`) {
        if (isEmptyOrNullOrUndefined(output_value)) return;
        
        let values_key = '';
        if (isNaN(Number(key))) {
          values_key = key;
        } else {
            // 处理像 "100 200" 这种值，第一个数字作为key的情况
            let datas = output_value.toString().split(' ');
            if (!isNaN(Number(datas[0]))) {
                 // 纯数字值，使用默认key
                 values_key = `var_${varIndex++}`;
            } else {
                 values_key = datas[0];
                 output_value = output_value.replace(`${datas[0]} `, '');
            }
        }

        // 技能本地化
        if (key == '#ValuesLoc') {
          if (!isEmptyOrNullOrUndefined(output_value) && output_value.trim() !== ``) {
             locAbilitySpecial = output_value;
          }
          return;
        }

        if (locAbilitySpecial != null) {
          let locKey = `dota_tooltip_ability_${main_key}_${values_key}`;
          locTokens.push({ key: locKey, value: locAbilitySpecial });
          locAbilitySpecial = null;
        }

        if (output_value != null && output_value.toString().trimStart().startsWith('{')) {
          contentLines.push(`${innerIndent}"${values_key}" ${output_value}`);
        } else {
          contentLines.push(`${innerIndent}"${values_key}" "${output_value}"`);
        }
        return;
      }

      // 3. 普通键值对
      // 处理本地化标记 #Loc
      if (key.includes('#Loc')) {
          if (!isEmptyOrNullOrUndefined(output_value) && output_value.trim() !== ``) {
             let locKey = key.replace('#Loc', ``).replace(`{}`, main_key);
             locTokens.push({ key: locKey, value: output_value });
          }
          return; 
      }

      output_value = deal_with_kv_value(output_value);
      contentLines.push(`${innerIndent}"${key}" "${output_value}"`);
    });

    // 2. 组装最终字符串：主键 { 内容 }
    // 注意：这里不再在循环里生成结尾的 }，而是统一在这里生成
    return `${baseIndent}"${main_key}" {\n${contentLines.join('\n')}\n${baseIndent}}`;
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

      workbook.forEach((sheet) => {
        let sheet_name = sheet.name;
        
        if (new RegExp(sheetsIgnore).test(sheet_name)) {
          console.log(cli.red(`${PLUGIN_NAME} Ignoring sheet ${sheet_name}...`));
          return;
        }

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
        
        // 简单KV处理 (仅两列)
        if (isSimpleKV(key_row) && autoSimpleKV) {
          const kv_data_simple = kv_data.map((row) => {
            return `\t"${row[0]}" "${row[1]}"`;
          });
          kv_data_str = `${kv_data_simple.join('\n')}`;
        } else {
          // 复杂KV处理
          const kv_data_complex = kv_data.map((row) => {
            if (isEmptyOrNullOrUndefined(row[0])) return;
            return convert_row_to_kv(row, key_row);
          });
          // 过滤掉空行再拼接
          kv_data_str = kv_data_complex.filter(x => x).join('\n');
        }

        mergedKVContent += `\n// --- Sheet: ${sheet_name} ---\n${kv_data_str}\n`; 
      });

      // 文件名逻辑
      let finalFilename: string;
      if (outputFilename) {
        finalFilename = outputFilename;
      } else if (firstSheetName) {
        finalFilename = firstSheetName;
      } else {
        finalFilename = path.basename(file.path, path.extname(file.path)) + '_merged';
      }
      const outputBasename = `${finalFilename}${kvFileExt}`;

      if (mergedKVContent.trim() !== '') {
        const out_put = `// this file is auto-generated by Xavier's sheet_to_kv 
// Source: ${file.basename}

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