@echo off
chcp 65001 >nul
echo ========================================================
echo   E-Office Editor - กำลังเปิดเซิร์ฟเวอร์ในเครื่อง...
echo   URL: http://localhost:8080
echo ========================================================
echo.

where python >nul 2>nul
if %errorlevel% equ 0 (
    start http://localhost:8080
    python -m http.server 8080 --directory dist
    goto end
)

where node >nul 2>nul
if %errorlevel% equ 0 (
    start http://localhost:8080
    node -e "const http=require('http'),fs=require('fs'),path=require('path');const dir=path.resolve('dist');const mime={'.html':'text/html','.js':'text/javascript','.css':'text/css','.png':'image/png','.jpg':'image/jpeg'};http.createServer((req,res)=>{let f=path.join(dir,req.url==='/'?'index.html':req.url.split('?')[0]);if(!fs.existsSync(f))f=path.join(dir,'index.html');const ext=path.extname(f);res.writeHead(200,{'Content-Type':mime[ext]||'application/octet-stream'});fs.createReadStream(f).pipe(res);}).listen(8080,()=>console.log('Server running on http://localhost:8080'));"
    goto end
)

echo [คำเตือน] ไม่พบ Python หรือ Node.js ในเครื่อง
pause

:end
