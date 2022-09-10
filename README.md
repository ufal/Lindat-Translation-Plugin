Plugin for MS Word.

Dev install and run:
--------------------
Powershell
```
git clone https://github.com/ufal/Lindat-Translation-Plugin.git
cd '.\Lindat-Translation-Plugin\'
# taskkill /IM msedgewebview2.exe /F # may help if previous instance was not closed properly
npm i
npm start
```

Now you should see opened Word document with button at right up.
After you press it, Lindat plugin window will open.
From here you can try various buttons.