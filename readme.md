# 打包命令
pyinstaller -c --clean --icon=app.ico -F -w  --add-data "C:\\Windows\\System32\\downlevel\\*;." .\ExcelSplitter.py --hidden-import=PyQt5.sip

# 重要提示
qypt==5.8.2
python==3.6.10
太高版本不支持windows7

打包时需要将system32\downlevel下的dll全部打包，否则win7环境下运行时会报错

# gitignore更新之后
git rm -r --cached .
git add .
git commit -m "update gitignore"