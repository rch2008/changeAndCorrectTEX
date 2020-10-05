# tex修正，试题分割
docx转换LaTeX，修正，分割工具

使用docx2tex进行转换 https://github.com/transpect/docx2tex

CopyAndDel.bat和convert.bat要放在docx2tex目录中

使用https://github.com/Kr00l/VBCCR 项目中的RichTextBox，VBCCR16.OCX可能需要自己注册。

docx转换LaTeX，修正，分割工具

使用docx2tex进行转换 https://github.com/transpect/docx2tex

CopyAndDel.bat和convert.bat要放在docx2tex目录中

使用https://github.com/Kr00l/VBCCR 项目中的RichTextBox，VBCCR16.OCX可能需要自己注册。

win7/win8/win10问题所在：

 64位系统一般都是可以安装32位程序的，只是需要执行 C:\Windows\SysWOW64\regsvr32.exe；

而不是  C:\Windows\System32\regsvr32.exe 。

【解决方法】 

把***.ocx拷贝到系统“C:\Windows\SysWOW64”文件夹下；

以管理员身份运行“C:\Windows\SysWOW64”文件夹下的“cmd.exe”；  

执行 regsvr32 ***.ocx即可注册成功。

若是32位系统，则把“C:\Windows\SysWOW64”替换为“C:\Windows\System32”，步骤相同

2020-08-31  包含以下更改
1. 选择题选项分割
2. 待续

2020-10-05

增加学生版试题答案拼接为教师版
试题答案分界为“\n参考答案”

单击join拼接结果查看，Ctrl+Join对刚查看的文件进行拼接