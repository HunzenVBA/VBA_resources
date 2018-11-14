UserForm3.bxDepartment.Value = Worksheets(1).ComboBox3.Value
UserForm3.bxSchlag1.Value = Worksheets(1).ComboBox1.Value
UserForm3.bxSchlag2.Value = Worksheets(1).ComboBox2.Value

UserForm3.bxDepartment.List = Worksheets(1).ComboBox3.List
UserForm3.bxSchlag1.List = Worksheets(1).ComboBox1.List
UserForm3.bxSchlag2.List = Worksheets(1).ComboBox2.List

bxDepartment_Change = ComboBox3_Change
UserForm3.bxDepartment_Change = ComboBox3_Change

bxDepartment_Change() = ComboBox3_Change
bxSchlag1_Change() = ComboBox1_Change
bxSchlag2_Change() = ComboBox2_Change

UserForm3.txtLogin.Value = Worksheets(1).Cells(4, 4).Value
UserForm3.txtDate.Value = Worksheets(1).Cells(4, 3).Value
UserForm3.txtOk.Value = Worksheets(1).Cells(4, 5).Value


UserForm3.txtComment.Value = UserForm1.TextBox1


Parcel.DHL - 123 - 321
