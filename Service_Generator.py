import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import time
import sys
import subprocess
import psutil

class MyService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'MyService'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_alive = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def is_process_running(self,process_name):
        for process in psutil.process_iter(['pid', 'name']):
            if process.info['name'] == process_name:
                return True
        return False

    def service(self):
        exe_loc=sys.argv[1]
        subprocess.run([exe_loc])

    def main(self):
        k=True
        while self.is_alive:
            if k:
                exe_loc=sys.argv[1]
                subprocess.run([exe_loc])
                k=False
            else:
                process_name=sys.argv[1].split('//')[-1]
                if not self.is_process_running(process_name):
                    exe_loc=sys.argv[1]
                    subprocess.run([exe_loc])
            time.sleep(1)

if __name__ == '__main__':
    if len(sys.argv) == 2:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(MyService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(MyService)