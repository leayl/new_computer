import wx


class SetControls:
    def __init__(self, win):
        """
        为窗口设置控件
        :param win:
        :return:
        """
        self.file_path_text = wx.TextCtrl(win, pos=(5, 5), size=(260, 30))
        self.open_file_btn = wx.Button(win, label="打开", pos=(280, 5), size=(50, 30))
        self.save_file_btn = wx.Button(win, label="保存", pos=(345, 5), size=(50, 30))
        self.content_text = wx.TextCtrl(win, pos=(5, 40), size=(390, 255), style=wx.TE_MULTILINE)

        self.open_file_btn.Bind(wx.EVT_BUTTON, self.open_file)

    def open_file(self, event):
        file_path = self.file_path_text.GetValue()
        with open(file_path, 'r', encoding="utf8") as f:
            self.content_text.SetValue(f.read())


if __name__ == '__main__':
    app = wx.App()
    win = wx.Frame(None, title='哇哦', size=(400, 300))
    SetControls(win)
    win.Show()
    app.MainLoop()
