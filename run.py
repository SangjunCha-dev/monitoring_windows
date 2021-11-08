import platform
import datetime as dt
import time

from monitoring_windows import SystemInfo, SystemMonitoring


class Main:
    @staticmethod
    def start():
        system_info = SystemInfo()
        system_monitoring = SystemMonitoring()

        info = system_info.run()
        print(f"시스템 정보 = {info}")

        while 1:
            if dt.datetime.now().strftime('%S') == '00':
                usage = system_monitoring.run()
                print(f"시스템 사용량 = {usage}")
            time.sleep(1)

if __name__ == '__main__':
    if platform.system() == 'Windows':
        Main.start()
    else:
        print('windows 모니터링만 지원합니다.')
