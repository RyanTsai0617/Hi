"# Hi" 

## sreits_xlsx_fix.py — S-REITs xlsx OneDrive/Excel 相容性修復

openpyxl 產出的 S-REITs 週報 xlsx 預設使用 inline string（無共用字串表），在
**OneDrive / Excel 網頁版**上常被判定「無法開啟活頁簿 / 內容需修復」。這支工具把
檔案原地轉成 **shared string** 格式（與可正常開啟的版本同款），只改字串儲存方式，
資料 / 超連結 / 合併 / 欄寬 / 樣式 100% 保留。

```bash
# 原地修復（覆蓋原檔），可多檔
python sreits_xlsx_fix.py 新加坡REITs_新聞_YYYYMMDD_含摘要.xlsx *.xlsx
```

建議接在產表流程每個 `wb.save(out_xlsx)` 之後呼叫，讓交付主檔本身就能開，不需再另存
「可開啟修復版」：

```python
from sreits_xlsx_fix import fix_xlsx
wb.save(out_xlsx)
fix_xlsx(out_xlsx)   # 原地轉 shared string
```

