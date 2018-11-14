Option Explicit

Sub test()

UserForm3.bxDepartment.Value = UserForm3.bxDepartment.Value
UserForm3.bxSchlag1.Value = UserForm3.bxSchlag1.Value
UserForm3.bxSchlag2.Value = UserForm3.bxSchlag2.Value

UserForm3.bxDepartment.List = UserForm3.bxDepartment.List
UserForm3.bxSchlag1.List = UserForm3.bxSchlag1.List
UserForm3.bxSchlag2.List = UserForm3.bxSchlag2.List

bxDepartment_Change = bxDepartment_Change()
UserForm3.bxDepartment_Change = bxDepartment_Change()

bxDepartment_Change() = bxDepartment_Change()
bxSchlag1_Change() = bxSchlag1_Change()
bxSchlag2_Change() = bxSchlag2_Change()

UserForm3.txtLogin.Value = UserForm3.txtLogin.Value
UserForm3.txtDate.Value = UserForm3.txtDate.Value
UserForm3.txtOk.Value = Worksheets(1).Cells(4, 5).Value

btnComment_Click = btnComment_Click

UserForm3.txtComment.Value = UserForm1.TextBox1

End Sub
