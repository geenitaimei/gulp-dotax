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
    indent = '    ',
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

  function convert_row_to_kv(row: string[], key_row: string[]): string {
    let main_key = row[0];

    if (typeof main_key == 'string' && main_key.trim() !== main_key) {
        console.warn(cli.red(`${main_key} 前后有空格，请检查！`));
    }

    let attachWearablesBlock = false;
    let abilityValuesBlock = false;
    let varIndex = 0;
    let locAbilitySpecial = null;

    const baseIndent = indent || '\t';
    const innerIndent = baseIndent + baseIndent;

    // --- 修改点 1: 增加 AbilityValues 内容缓存 ---
    let abilityValuesContent: string[] = [];

    let contentLines: string[] = [];

    key_row.forEach((key, i) => {
      if (isEmptyOrNullOrUndefined(key)) return;
      if (i === 0) return;

      let output_value = row[i];

      // --- 状态机逻辑控制 ---
      if (key === `AttachWearables[{]`) attachWearablesBlock = true;
      if (attachWearablesBlock && key == `}]`) attachWearablesBlock = false;

      // 开启 AbilityValues 模式
      if (key === `AbilityValues[{]`) {
          abilityValuesBlock = true;
          abilityValuesContent = []; // 重置缓存
          return; // 这一列不输出，直接跳过
      }
      // 关闭 AbilityValues 模式
      if (abilityValuesBlock && key === `}]`) {
          abilityValuesBlock = false;
          // 模式结束时，生成最终的 KV 块
          if (abilityValuesContent.length > 0) {
              contentLines.push(`${innerIndent}"AbilityValues" { ${abilityValuesContent.join(" ")} }`);
          }
          return;
      }

      // --- 块内特殊处理 ---

      // 1. 饰品处理 (原有逻辑)
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

      // 2. AbilityValues 内部数据收集
      if (abilityValuesBlock) {
        // 跳过空行
        if (isEmptyOrNullOrUndefined(output_value)) return;

        let values_key = '';
        // 如果 Key 是数字 (如 1, 2)，通常表示该行的 Value 第一个词是 Key
        if (isNaN(Number(key))) {
          values_key = key;
        } else {
            let datas = output_value.toString().split(' ');
            if (!isNaN(Number(datas[0]))) {
                 values_key = `var_${varIndex++}`;
            } else {
                 values_key = datas[0];
                 output_value = output_value.replace(`${datas[0]} `, '');
            }
        }

        // 本地化标记处理
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

        // 将解析后的键值对存入缓存，格式: "key" "value"
        if (output_value != null && output_value.toString().trimStart().startsWith('{')) {
          abilityValuesContent.push(`"${values_key}" ${output_value}`);
        } else {
          abilityValuesContent.push(`"${values_key}" "${output_value}"`);
        }
        return;
      }

      // 3. 普通键值对处理 (原有逻辑)
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

    // 组装最终字符串
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
          kv_data_str = kv_data_complex.filter(x => x).join('\n');
        }

        mergedKVContent += `\n// --- Sheet: ${sheet_name} ---\n${kv_data_str}\n`;
      });

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