# SPO365-JavaScript
## Quick and simple JavaScript on SPO365 {SharePoint Online}

Kadang-kadang, kita ingin sekali melakukan testing suatu code JavaSript ke Site berbasis SharePoint Online. Tetapi kita menghadapi beberapa kendala:
1. kita tidak bisa melakukan Develop WebPart menggunakan SharePoint Framework SPFX
2. kita tidak punya akses untuk melakukan upload file *.sppkg ke SharePoint Online
3. kita tidak punya Script Editor {Modern} di SharePoint Online

Semua point-point di atas, saya rangkum sebagai berikut:
<img src="images/JavaScript%20to%20SharePoint%20Online/TaskLists.png"/>

Kita bisa menggunakan file **01 Quick JavaScript to SharePoint Online.js** yang dieksekusi pada Console {Browser - Developer Tools}, untuk testing code JavaScript yang kita inginkan:

<img src="images/JavaScript%20to%20SharePoint%20Online/Console.png"/>

Selama eksekusi code JavaScript di Console, Site berbasis SharePoint Online akan tampil seperti berikut:

<img src="images/JavaScript%20to%20SharePoint%20Online/SharePoint-Site.png"/>

hasil eksekusi code JavaScript, beberapa file berhasil diupload ke SharePoint Online:

<img src="images/JavaScript%20to%20SharePoint%20Online/Upload-files.png"/>
