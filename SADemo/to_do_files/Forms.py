import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows import Forms
from System.Drawing import Font

font = Font("Microsoft Sans Serif", 10)

class MyDialog(Forms.Form):
    def __init__(self):
        self.Text ='Prompt'
        self.Height = 120
        self.Width = 400
        self.StartPosition = Forms.FormStartPosition.CenterScreen
        lbl_prompt = Forms.Label(Visible=True, Left=0, Top=12, Text="Answer", Width=150, Font=font)
        txt_input = Forms.TextBox(Visible=True, Left=160, Top=10, Width=180, Text="", Font=font)
        ok_button = Forms.Button(Text='OK', Visible=True, Left=50, Top=50)
        self.AcceptButton = ok_button
        ok_button.Click += self.OnOK
        cancel_button = Forms.Button(Text='Cancel', Visible=True, Left=150, Top=50)
        cancel_button.Click += self.OnCancel
        self.Controls.Add(lbl_prompt)
        self.Controls.Add(txt_input)
        self.Controls.Add(ok_button)
        self.Controls.Add(cancel_button)

    def OnOK(self, *args):
        self.result = self.Controls[1].Text
        self.DialogResult = Forms.DialogResult.OK
        self.Close()
    def OnCancel(self, *args):
        self.result = "Spam"
        self.DialogResult = Forms.DialogResult.Cancel
        self.Close()

def user_query(prompt):
    md = MyDialog()
    md.Controls[0].Text = prompt + ":"
    if md.ShowDialog() == Forms.DialogResult.OK:
        return md.result
    else:
        return ""

#md = MyDialog()
#md.ShowDialog()