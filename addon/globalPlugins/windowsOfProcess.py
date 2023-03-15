#-*- coding:utf-8 -*-

import globalPluginHandler,api,winUser,appModuleHandler,wx,gui,time,ui,os,winKernel,winsound
import addonHandler
from ctypes import WINFUNCTYPE,c_int,c_void_p
from tones import beep

addonHandler.initTranslation()

pl=[]
pi=0
hd={}
index=0
pid=None
old_pid=None
fg=None
prev_time=0
iwl=None
epl=[]

def play(string):
	path= os.path.join(os.path.dirname(__file__),string+".wav")
	winsound.PlaySound(path, winsound.SND_FILENAME|winsound.SND_ASYNC)

Callback=WINFUNCTYPE(c_int,c_int,c_void_p)

@Callback
def callback(hwnd,pid):
	global pl,hd
	if winUser.isWindowVisible(hwnd) and winUser.isWindowEnabled(hwnd):
		wid=winUser.getWindowThreadProcessID(hwnd)[0]
		processName = appModuleHandler.getAppNameFromProcessID(wid, True)
		if not processName in epl and not wid in pl:
			pl.append(wid)
			hd[wid]=[]
		if processName=="explorer.exe":
			if winUser.getWindowText(hwnd) and not hwnd in hd[wid]:
				hd[wid].append(hwnd)
			return True
		if wid in pl and not hwnd in hd[wid]:
			hd[wid].append(hwnd)
	return True

def switchWindow():
	global pid,index,old_pid,fg,prev_time
	fg=api.getForegroundObject()
	pid=fg.processID
	if old_pid != pid:
		old_pid = pid
		index=0
	cur_time=time.time()
	if cur_time - prev_time > 2:
		prev_time=cur_time
		winUser.user32.EnumWindows (callback,pid)
	nextWindow()

def nextWindow():
	global index,fg
	if len(hd[pid]) > 1:
		index=(index+1)%len(hd[pid])
		if fg.windowHandle == hd[pid][index]:
			return nextWindow()
		if winUser.isWindowVisible(hd[pid][index]) and winUser.isWindowEnabled(hd[pid][index]) and winUser.user32.SetForegroundWindow(hd[pid][index]):
			showWindow()
		else:
			hd[pid].remove(hd[pid][index])
			nextWindow()
	else:
		beep(500,30)

def switchProcess():
	global pid,fg,prev_time
	fg=api.getForegroundObject()
	pid=fg.processID
	cur_time=time.time()
	if cur_time - prev_time > 2:
		prev_time=cur_time
		winUser.user32.EnumWindows (callback,pid)
	nextProcess()

def nextProcess():
	global pi
	if len(pl)>1:
		pi=(pi+1)%len(pl)
		if pid==pl[pi]:
			return nextProcess()
		if not nextProcessWindow(pl[pi]):
			pl.remove(pl[pi])
			nextProcess()
	else:
		beep(600,30)

def nextProcessWindow(newPid):
	try:
		if hd[newPid][0]:
			if winUser.isWindowVisible(hd[newPid][0]) and winUser.isWindowEnabled(hd[newPid][0]) and winUser.user32.SetForegroundWindow(hd[newPid][0]):
				showWindow()
				return True
			else:
				hd[newPid].remove(hd[newPid][0])
				nextProcessWindow(newPid)
		else:
			beep(500,30)
			return False
	except:
		return False

def showWindow():
		wx.CallAfter(winUser.keybd_event,92, 0, 1, 0)
		wx.CallAfter(winUser.keybd_event,38, 0, 1, 0)
		wx.CallAfter(winUser.keybd_event,38, 0, 3, 0)
		wx.CallAfter(winUser.keybd_event,92, 0, 3, 0)


def showWindowsList():
	global iwl
	if iwl:
		return
	fg=api.getForegroundObject()
	winUser.user32.EnumWindows (callback,pid)
	hl=[]
	try:
		for p in hd:
			for w in hd[p]:
				if winUser.isWindowVisible(w) and winUser.isWindowEnabled(w):
					hl.append(w)
				else:
					hd[p].remove(w)
			if len(hd[p])==0:
				del hd[p]
				pl.remove(p)
	except:
		pass
	title=_("{} window， {} process").format(len(hl),len(pl))
	try:
		selection=hl.index(fg.windowHandle)
	except:
		selection=0
	iwl=windowsListView(title,hl,selection)

class windowsListView(wx.Dialog):
	def __init__(self,title,hl,selection):
		super(windowsListView,self).__init__(gui.mainFrame,title=title)
		listBoxSizer = wx.BoxSizer(wx.VERTICAL)
		self.st = wx.StaticText(self,-1,_("Window"))
		listBoxSizer.Add(self.st,0.5, wx.ALL, 10)
		self.listBox = wx.ListBox(self,-1)
		listBoxSizer.Add(self.listBox,0,wx.ALL| wx.EXPAND,10)
		self.SetSizer(listBoxSizer)
		listBoxSizer.Fit(self)
		self.hl=hl
		for w in self.hl:
			self.listBox.Append(winUser.getWindowText(w))
		buttonsSizer = wx.BoxSizer(wx.VERTICAL)
		b_killProcess = wx.Button(self, -1,_("&Kill the window of process"))
		buttonsSizer.Add(b_killProcess ,0, wx.ALL| wx.CENTER| wx.EXPAND,10)
		buttonsSizer.Add(self.CreateButtonSizer(wx.OK), wx.ALL| wx.CENTER|wx.EXPAND,10)
		buttonsSizer.Add(self.CreateButtonSizer(wx.CANCEL), wx.ALL| wx.CENTER|wx.EXPAND,10)
		self.SetSizer(buttonsSizer)
		buttonsSizer.Fit(self)
		b_killProcess .Bind(wx.EVT_BUTTON, self.onKillProcess )
		self.Bind(wx.EVT_BUTTON, self.onOk, id=wx.ID_OK)
		self.Bind(wx.EVT_BUTTON, self.onCancel, id=wx.ID_CANCEL)
		self.listBox.SetFocus()
		self.listBox.SetSelection(selection)
		self.CenterOnScreen()
		self.Raise()
		self.Maximize()
		self.Show()

	def onKillProcess (self, event):
		wid=winUser.getWindowThreadProcessID(self.hl[self.listBox.GetSelection()])[0]
		handle = winKernel.kernel32.OpenProcess(1, 0, wid)
		res = winKernel.kernel32.TerminateProcess(handle, 0)
		winKernel.kernel32.CloseHandle(handle)
		if res:
			play("success")
			self.listBox.Clear()
			self.hl=[]
			global hd
			del hd[wid]
			try:
				for p in hd:
					for w in hd[p]:
						if winUser.isWindowVisible(w) and winUser.isWindowEnabled(w):
							self.hl.append(w)
							self.listBox.Append(winUser.getWindowText(w))
						else:
							hd[p].remove(w)
					if len(hd[p])==0:
						del hd[p]
						pl.remove(p)
			except:
				pass
			self.listBox.SetFocus()
			self.listBox.SetSelection(0)
		else:
			play("fail")

	def onOk(self, event):
		winUser.user32.SetForegroundWindow(self.hl[self.listBox.GetSelection()])
		showWindow()
		global iwl
		iwl=None
		self.Close()

	def onCancel(self,evt):
		global iwl
		iwl=None
		self.Destroy()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("WindowsOfProcess ")

	def __init__(self):
		super(GlobalPlugin, self).__init__()
		path = os.path.join(os.path.dirname(__file__),"epl.txt")
		with open(path,"r",encoding="utf-8") as f:
			global epl
			epl=f.read().split("\n")

	def script_switchWindow(self, gesture):
		switchWindow()
	script_switchWindow.__doc__ = _("switch between windows with the same process")

	def script_switchProcess(self, gesture):
		switchProcess()
	script_switchProcess.__doc__ = _("Switch windows of different process")

	def script_showWindowsList(self, gesture):
		showWindowsList()
	script_showWindowsList.__doc__ = _("Show window list")

	__gestures = {
"kb:control+.":"switchWindow",
"kb:control+shift+.":"switchProcess",
"kb:control+windows+.":"showWindowsList",
	}