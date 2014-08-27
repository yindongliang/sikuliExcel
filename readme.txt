技术点：读取excel文件，调用python脚本来实现对sikuli的使用
优点:安装环境配置方便，数据和逻辑有效的分离。
缺点：写excel的同时需要注意python的语法和sikuli的接口。



使用
.\jre\bin\java.exe  -jar excelsikuli.jar sikuli_IF.xlsx 将执行sikuli_IF.xlsx里面的所有的sheet
.\jre\bin\java.exe  -jar excelsikuli.jar sikuli_IF.xlsx sheet1将执行sikuli_IF.xlsx里面的sheet1
.\jre\bin\java.exe  -jar excelsikuli.jar sikuli_IF.xlsx sheet1,sheet2将执行sikuli_IF.xlsx里面的sheet1和sheet2

python的语法和sikuli的接口
请看sikuli_IF.xlsx的excel comment
http://doc.sikuli.org/
https://github.com/sikuli/sikuli
http://www.w3cschool.cc/python/python-tutorial.html


├─bin                         //vnc的执行程序目录
├─image                       //查找图片存放的地方
│  ├─com                      
│  ├─user                     //用户自己抓的图片
│  │  └─dereck

├─img                        //屏幕截图
├─jre                        //jre环境
│  ├─bin
│  │  ├─client
│  │  ├─dtplugin
│  │  ├─plugin2
│  │  └─server
│  └─lib                     
│      ├─applet
│      ├─cmm
│      ├─deploy
│      │  └─jqs
│      ├─ext
│      ├─fonts
│      ├─i386
│      ├─images
│      │  └─cursors
│      ├─management
│      ├─security
│      ├─servicetag
│      └─zi
│          ├─Africa
│          ├─America
│          │  ├─Argentina
│          │  ├─Indiana
│          │  ├─Kentucky
│          │  └─North_Dakota
│          ├─Antarctica
│          ├─Asia
│          ├─Atlantic
│          ├─Australia
│          ├─Etc
│          ├─Europe
│          ├─Indian
│          ├─Pacific
│          └─SystemV
├─lib                           //java sikuli lib包
│  ├─Microsoft.VC90.CRT
│  └─Microsoft.VC90.OPENMP
└─logs                         //log日志
