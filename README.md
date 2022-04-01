@[TOC](文章目录)
## 介绍
```
制作：Morgan   
工具：VisualBasic6.0
邮箱：MorganFish0508@163.com
     1037502595@qq.com
GitHub：https://github.com/MorganNotFound
CSDN:https://blog.csdn.net/MorganFish
欢迎点赞+收藏+下载+评论   
```
***附代码，附源程序，附成品***
## 代码&制作
这个程序说难不难，就是制作要有耐心，否则难以完成。不过这也是优点，可以尽情创作、更改，在总体的框架结构下，可以一步步的完善，下面我将展示一下我三个阶段的代码，并展示我的制作方法（但准确来说如果要改进还可以有v4.0、v5.0版本）
### v1.0 
    （一）新建一个标准exe并适当调大窗口，为画图提供空间   
    （二）插入三个HScrollBar控件即水平滚动条,并使用三个Label来对滚动条控制对象进行标注，右侧再加入一个垂直滚动条VScrollBar，更改HScroll1~3的Min属性为0，Max为255，VScroll1的Min为1，Max为15   
    （三）插入一个Label，删除其Caption   
    （四）插入代码
![新建](https://img-blog.csdnimg.cn/b343c18729f448a8b4c8dc3c782ffb5c.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBATW9yZ2FuRmlzaA==,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![插入控件](https://img-blog.csdnimg.cn/771b07fe3c1e40df96ebedd3d622a797.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBATW9yZ2FuRmlzaA==,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![插入代码](https://img-blog.csdnimg.cn/8714bc7ce2e443c08adc4a13e23d454b.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBATW9yZ2FuRmlzaA==,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
代码如下：
```
Private Sub Form_Load()
HScroll1.Min = 0
HScroll1.Max = 255
HScroll2.Min = 0
HScroll2.Max = 255
HScroll3.Min = 0
HScroll3.Max = 255
VScroll1.Min = 1
VScroll1.Max = 15
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.CurrentX = X
Form1.CurrentY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Line -(X, Y), RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub HScroll1_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub HScroll2_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub HScroll3_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub VScroll1_Change()
Form1.DrawWidth = VScroll1.Value
End Sub
```
这里是运用了MouseDown和MouseMove的检测，记录按下鼠标时的坐标，然后使用`line`进行绘画   
![在这里插入图片描述](https://img-blog.csdnimg.cn/5ca7145b41f3424bbebab927c506aa8f.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBATW9yZ2FuRmlzaA==,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)

### v2.0
***使用一个line1替代了变换颜色的label，并可以预览画笔粗细，新增TextBox显示RGB色号***   
先看看效果：   
![v2.0](https://img-blog.csdnimg.cn/aac49aa685f9425397a545b75153ca66.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBATW9yZ2FuRmlzaA==,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)

我插入了许多色块（label5(index as integer)）,单击选中绘画颜色，成品都在我的GitHub仓库，代码和v3.0的合并在了一起，由于3的代码其实不太完善，我都使用`'`把代码禁用了，可以自己更改哦~
### v3.0
***新增自定义颜色存储并支持画圆（增用了一个Combo1），并升级了TextBox使其同时可以控制颜色，即可以输入来改变颜色***   
看看效果：   
![v3.0](https://img-blog.csdnimg.cn/dcfa772099a34b5190625cd36cb606de.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBATW9yZ2FuRmlzaA==,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)

再看看代码：
```
Private Sub Command1_Click()
Form1.Cls
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
HScroll1.Min = 0
HScroll1.Max = 255
HScroll2.Min = 0
HScroll2.Max = 255
HScroll3.Min = 0
HScroll3.Max = 255
HScroll4.Min = 1
HScroll4.Max = 20
Combo1.AddItem "line"
'Combo1.AddItem "circle"
'Combo1.AddItem "straight line"
Label8.Visible = False
HScroll5.Visible = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.CurrentX = X
Form1.CurrentY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Combo1.Text = "line" Then
Label8.Visible = False
HScroll5.Visible = False
If Button = 1 Then Line -(X, Y), Line1.BorderColor
End If
'If Combo1.Text = "circle" Then
'Label8.Visible = True
'HScroll5.Visible = True
'If Button = 1 Then Circle (X, Y), ((M - X) ^ 2 + (N - Y) ^ 2) ^ (1 / 2), Line1.BorderColor
'End If
'If Combo1.Text = "line" Then
'Label8.Visible = False
'HScroll5.Visible = False
'If Button = 1 Then Line Step(X, Y)-Step(M, N), Line1.BorderColor
'End If
End Sub
'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form1.CurrentX = M
'Form1.CurrentY = N
'End Sub
Private Sub HScroll1_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = HScroll1.Value
End Sub
Private Sub HScroll2_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text2.Text = HScroll2.Value
End Sub
Private Sub HScroll3_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text3.Text = HScroll3.Value
End Sub
Private Sub HScroll4_Change()
Form1.DrawWidth = HScroll4.Value
Line1.BorderWidth = HScroll4.Value
Text4.Text = HScroll4.Value
End Sub
Private Sub Label5_Click(Index As Integer)
Line1.BorderColor = Label5(Index).BackColor
End Sub
Private Sub label6_click()
Line1.BorderColor = Label6.BackColor
End Sub
Private Sub Label7_Click()
Label6.BackColor = Line1.BorderColor
End Sub
Private Sub Text1_Change()
HScroll1.Value = Text1.Text
End Sub
Private Sub Text2_Change()
HScroll2.Value = Text2.Text
End Sub
Private Sub Text3_Change()
HScroll3.Value = Text3.Text
End Sub
Private Sub Text4_Change()
HScroll4.Value = Text4.Text
End Sub
```
我原想添加画直线与圆的代码，结果发现并不令人满意，就禁用了hhh...不过基本算是完成了目标，并适当的美化了程序
### v3.1
嘿嘿，为什么要有v3.1咧？是因为我闲嘛……不不不当然不是，这很重要，众所周知，vb只要类型不匹配就退出，万一输数据的时候一手误就退出咋办咧？别急别急，使用if判断一下不就好了嘛！先看看代码：
```
Private Sub Command1_Click()
Form1.Cls
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
HScroll1.Min = 0
HScroll1.Max = 255
HScroll2.Min = 0
HScroll2.Max = 255
HScroll3.Min = 0
HScroll3.Max = 255
HScroll4.Min = 1
HScroll4.Max = 20
Combo1.AddItem "line"
'Combo1.AddItem "circle"
'Combo1.AddItem "straight line"
Label8.Visible = False
HScroll5.Visible = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.CurrentX = X
Form1.CurrentY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Combo1.Text = "line" Then
Label8.Visible = False
HScroll5.Visible = False
If Button = 1 Then Line -(X, Y), Line1.BorderColor
End If
'If Combo1.Text = "circle" Then
'Label8.Visible = True
'HScroll5.Visible = True
'If Button = 1 Then Circle (X, Y), ((M - X) ^ 2 + (N - Y) ^ 2) ^ (1 / 2), Line1.BorderColor
'End If
'If Combo1.Text = "line" Then
'Label8.Visible = False
'HScroll5.Visible = False
'If Button = 1 Then Line Step(X, Y)-Step(M, N), Line1.BorderColor
'End If
End Sub
'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form1.CurrentX = M
'Form1.CurrentY = N
'End Sub
Private Sub HScroll1_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = HScroll1.Value
End Sub
Private Sub HScroll2_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text2.Text = HScroll2.Value
End Sub
Private Sub HScroll3_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text3.Text = HScroll3.Value
End Sub
Private Sub HScroll4_Change()
Form1.DrawWidth = HScroll4.Value
Line1.BorderWidth = HScroll4.Value
Text4.Text = HScroll4.Value
End Sub
Private Sub Label5_Click(Index As Integer)
Line1.BorderColor = Label5(Index).BackColor
End Sub
Private Sub label6_click()
Line1.BorderColor = Label6.BackColor
End Sub
Private Sub Label7_Click()
Label6.BackColor = Line1.BorderColor
End Sub
Private Sub Text1_Change()
If Text1.Text = "" Then
Text1.Text = 0
End If
If Text1.Text > 255 Then
Text1.Text = 255
End If
HScroll1.Value = Text1.Text
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text2_Change()
If Text2.Text = "" Then
Text2.Text = 0
End If
If Text2.Text > 255 Then
Text2.Text = 255
End If
HScroll2.Value = Text2.Text
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text3_Change()
If Text3.Text = "" Then
Text3.Text = 0
End If
If Text3.Text > 255 Then
Text3.Text = 255
End If
HScroll3.Value = Text3.Text
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text4_Change()
If Text4 = "" Then
Text4.Text = 1
End If
If Text4.Text = 0 Then
Text4.Text = 1
End If
If Text4.Text > 20 Then
Text4.Text = 20
End If
HScroll4.Value = Text4.Text
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
```
这个通过禁用数字和退格之外的按键来确保不会输错，又通过将超出范围的数字给强行改回来确保不会在清空Textbox的时候报错，不理解的可以到网上某度一下
## 下载地址
[点此前往下载](https://github.com/MorganNotFound/vb_paint)   
作者将会努力更新，尽量做到高度还原，感谢大家的支持，也邀请大家一起贡献力量哦~
## 附言
其实还有许多美中不足之处，如果能想到更好的方案，烦请各位朋友在评论区指教，谢谢！   
***如果您喜欢我的作品，请点赞&关注哦，谢谢啦！***
