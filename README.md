# Excel 批量合并去重清洗工具

把一个文件夹里的几十个 Excel/CSV 表格，**一键合并成一张干净的总表**，自动去重、去空格、删空行，并生成一张"处理报告"告诉你每个文件贡献了多少行、去掉了多少重复。

> 适合：多门店/多渠道客户名单合并、多月订单汇总、多人填写的表格统一、爬取数据去重。

## 效果演示

输入：3 个门店的客户名单（共 13 行，手机号有 3 条跨表重复）
输出：合并后 10 行有效数据 + 一张处理报告 sheet

```
发现 3 个文件，开始合并……
  [OK] 门店A_客户名单.xlsx  5 行
  [OK] 门店B_客户名单.xlsx  4 行
  [OK] 门店C_客户名单.xlsx  4 行
按 ['手机号'] 去重：删除 3 条重复
完成 → merged_result.xlsx
最终 10 行（原 13 行，去重删 3 行）
```

## 功能

- 自动识别文件夹里所有 `.xlsx / .xls / .csv`
- CSV 自动尝试 `utf-8 / gbk` 编码，中文不乱码
- 所有列按文本读取，**手机号/身份证号不会变成科学计数法**
- 按指定列去重（如按"手机号"）或整行去重，可选保留第一条/最后一条
- 清洗：去首尾空格、删完全空白行
- 输出带"处理报告"sheet，可追溯每个来源文件

## 使用

```bash
pip install -r requirements.txt

# 按手机号去重
python merge_tool.py --input ./data --output merged.xlsx --dedup-cols 手机号

# 按多列去重
python merge_tool.py --input ./data --output merged.xlsx --dedup-cols 手机号,订单号

# 整行完全相同才算重复
python merge_tool.py --input ./data --output merged.xlsx --dedup-all

# 只合并+清洗，不去重
python merge_tool.py --input ./data --output merged.xlsx
```

先跑 `python generate_demo_data.py` 可生成演示数据体验效果。

## 打包成 exe（交付给不懂技术的客户）

```bash
pip install pyinstaller
pyinstaller -F -n Excel合并去重工具 merge_tool.py
```

客户拿到 `dist/Excel合并去重工具.exe` 双击即可用，无需安装 Python。

## 参数

| 参数 | 说明 | 默认 |
|---|---|---|
| `--input / -i` | 待合并表格所在文件夹 | 必填 |
| `--output / -o` | 输出文件名 | merged_result.xlsx |
| `--dedup-cols` | 按这些列去重，逗号分隔 | 空（不去重） |
| `--dedup-all` | 整行去重 | 关 |
| `--keep` | 重复时保留 first/last | first |

## License

MIT
