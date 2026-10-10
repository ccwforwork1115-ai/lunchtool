# GitHub 更新方式

本資料夾已連接 https://github.com/ccwforwork1115-ai/lunchtool 。

每次修改後，可請 Codex「檢查修改並更新到 GitHub」。

自行操作時，在此資料夾開啟 PowerShell：

```powershell
git status
git add .
git commit -m "說明這次修改"
git push
```

使用另一台電腦或在 GitHub 修改後，開始工作前先執行 `git pull --ff-only`。如遇到衝突或拒絕推送，請先處理差異，不要強制推送。

PDF、Excel、CSV、圖片、跨年度測試資料與暫存目錄已排除，不會隨一般更新上傳。程式內建的範本仍屬 HTML 的一部分。

儲存檔案只會更新本機；完成 commit 和 push 後才會更新 GitHub。
