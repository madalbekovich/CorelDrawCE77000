import win32serviceutil
import win32service
import win32event
import servicemanager
import time

class MyService(win32serviceutil.ServiceFramework):
    _svc_name_ = "MyPythonService"
    _svc_display_name_ = "My Python 24/7 Script"

    def init(self, args):
        win32serviceutil.ServiceFramework.init(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.running = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ""))
        self.main()

    def main(self):
        while self.running:

            time.sleep(5)

if __name__ == 'main':
    win32serviceutil.HandleCommandLine(MyService)