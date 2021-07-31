from System.IO import FileSystemWatcher, FileMode, FileAccess, FileShare
from System.Threading import Thread
import System.IO.FileInfo as s
import me_import as i
import os
import smtplib

watcher = FileSystemWatcher()
watcher.Path = r'G:\Electricity\ME\MISO\Financials\to_import'


def onChanged(source,event):
	print 'Changed:', event.ChangeType, event.FullPath
	files = [];
	for file in os.listdir(r'G:\Electricity\ME\MISO\Financials\to_import'):
		f = s(event.FullPath)
		try:
			z = f.Open(FileMode.Open, FileAccess.Read, FileShare.None)
			z.Close()
			files.append(file);
		except Exception, e:
			print Exception
			print e
			print "------------------------"	
	try:
		i.update(files);
	except Exception, e:
		print Exception
		print e
		i.email(Exception, e)
watcher.Created += onChanged
watcher.EnableRaisingEvents = True 
	
while(True):
	Thread.CurrentThread.Join(1000 * 60)