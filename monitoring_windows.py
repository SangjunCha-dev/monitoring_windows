import subprocess
import json


class PowerShell:
    '''
    powershell 명령어 실행
    '''
    @staticmethod
    def convert_to_json(cmd: str):
        result = None
        try:
            result = subprocess.Popen(f'powershell.exe {cmd} | ConvertTo-JSON', stdout=subprocess.PIPE)
            result = json.loads(result.stdout.read().decode('cp949'))
        except Exception as ex:
            print(f"[PowerShell] convert_to_json error : {str(ex)}")
        return result
    
    @staticmethod
    def isnumeric(cmd: str):
        result = subprocess.Popen(f'powershell.exe {cmd}', stdout=subprocess.PIPE)
        result = result.stdout.read().decode('cp949').strip()

        # 공백 or 에러를 반환하는 경우 예외처리
        if result.isnumeric(): 
            result = int(result)
        else:
            result = None

        return result


class SystemInfo:
    '''
    시스템 정보 조회
    '''
    def __init__(self):
        self.os_instance = {}
        self.cpu_instance = {}
        self.gpu_instance = {}
        self.memory_instance = {}
        self.disk_instance = {}
        self.volume_instance = {}
        self.device_info = {}

    def run(self):
        try:
            self.os_instance, memory_info = self.os_info()
            self.cpu_instance = self.cpu_info()
            self.gpu_instance = self.gpu_info()
            self.memory_instance = self.memory_info()
            self.memory_instance.update(memory_info)
            self.disk_instance = self.disk_info()
            self.volume_instance = self.volume_info()

            self.device_info = {
                'monitoring': 'info',
                'os': self.os_instance,
                'cpu': self.cpu_instance,
                'gpu': self.gpu_instance,
                'memory': self.memory_instance,
                'disk': self.disk_instance,
                'volume': self.volume_instance,
            }

            with open('./system_info.json', 'w') as json_file:
                json_file.write(json.dumps(self.device_info))

        except Exception as ex:
            print(f"[SystemInfo] run error : {str(ex)}")

        return self.device_info

    def os_info(self):
        '''
        운영체제 및 메모리 사용량 정보 조회
        '''
        try:
            ps_cmd = 'Get-CimInstance -Class Win32_OperatingSystem | Select-Object -Property Caption, OSArchitecture, Version, TotalVisibleMemorySize, FreePhysicalMemory'
            os_info = PowerShell.convert_to_json(cmd=ps_cmd)
            os_info['OSArchitecture'] = os_info['OSArchitecture'].replace('비트', 'bit')

            os_instance = {
                'name': os_info['Caption'],
                'os_architecture': os_info['OSArchitecture'],
                'version': os_info['Version'],
            }

            os_info['TotalVisibleMemorySize'] //= 1024  # MB
            os_info['FreePhysicalMemory'] //= 1024  # MB
            use_memory = (os_info['TotalVisibleMemorySize'] - os_info['FreePhysicalMemory'])
            use_memory_percent = round((use_memory/(os_info['TotalVisibleMemorySize']))*100, 1)

            memory_info = {
                'total_size': os_info['TotalVisibleMemorySize'],
                'use_size': os_info['TotalVisibleMemorySize'] - os_info['FreePhysicalMemory'],
                'use_percent': use_memory_percent
            }

        except Exception as ex:
            print(f"[SystemInfo] os_info error : {str(ex)}")

        return os_instance, memory_info

    def cpu_info(self):
        '''
        CPU 정보 조회
        '''
        try:
            ps_cmd = 'Get-CimInstance -ClassName Win32_Processor | Select-Object -Property Name, MaxClockSpeed, LoadPercentage'
            cpu_info = PowerShell.convert_to_json(cmd=ps_cmd)

            use_percent = 0
            cpu_instance = {'name': []}
            if isinstance(cpu_info, dict):
                cpu_name = f"{cpu_info['Name'].strip()} {round(cpu_info['MaxClockSpeed']/1000, 1)}GHz"
                cpu_instance['name'].append(cpu_name)

                cpu_instance['use_percent'] = cpu_info['LoadPercentage']
            else:
                for info in cpu_info:
                    cpu_name = f"{info['Name'].strip()} {round(info['MaxClockSpeed']/1000, 1)}GHz"
                    cpu_instance['name'].append(cpu_name)

                    use_percent += info['LoadPercentage']
                cpu_instance['use_percent'] = round(use_percent/len(cpu_info))

        except Exception as ex:
            print(f"[SystemInfo] cpu_info error : {str(ex)}")

        return cpu_instance

    def gpu_info(self):
        '''
        GPU 정보 조회
        '''
        try:
            ps_cmd = 'Get-CimInstance -ClassName  Win32_VideoController | Select-Object -Property Name, AdapterRAM'
            gpu_info = PowerShell.convert_to_json(cmd=ps_cmd)

            gpu_total_memory = 0
            gpu_instance = {'name': []}
            if isinstance(gpu_info, dict):
                # uint32 초과한 비디오 메모리는 아래의 명령으로 다시 조회한다.(uint64 데이터형)
                if gpu_info['AdapterRAM'] == 4293918720:
                    ps_cmd = "(Get-ItemProperty -Path 'HKLM:\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\\0*' -Name HardwareInformation.qwMemorySize -ErrorAction SilentlyContinue).'HardwareInformation.qwMemorySize'"
                    gpu_info['AdapterRAM'] = PowerShell.isnumeric(cmd=ps_cmd)

                gpu_info['AdapterRAM'] //= 1048576  # MB
                gpu_name = f"{gpu_info['Name']} {gpu_info['AdapterRAM']//1024}GB"
                gpu_instance['name'].append(gpu_name)

                gpu_total_memory += gpu_info['AdapterRAM']
            else:
                for info in gpu_info:
                    # uint32 초과한 비디오 메모리는 아래의 명령으로 다시 조회한다.(uint64 데이터형)
                    if info['AdapterRAM'] == 4293918720:
                        ps_cmd = "(Get-ItemProperty -Path 'HKLM:\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\\0*' -Name HardwareInformation.qwMemorySize -ErrorAction SilentlyContinue).'HardwareInformation.qwMemorySize'"
                        info['AdapterRAM'] = PowerShell.isnumeric(cmd=ps_cmd)

                    info['AdapterRAM'] //= 1048576  # MB
                    gpu_name = f"{info['Name']} {info['AdapterRAM']//1024}GB"
                    gpu_instance['name'].append(gpu_name)

                    gpu_total_memory += info['AdapterRAM']

            ps_cmd = "(((Get-Counter '\GPU Process Memory(*)\Local Usage').CounterSamples | where CookedValue).CookedValue | measure -sum).sum"
            gpu_use_memory = PowerShell.isnumeric(cmd=ps_cmd)

            if isinstance(gpu_use_memory, int):
                gpu_use_memory //= 1048576  # MB
                gpu_use_percent = round((gpu_use_memory / gpu_total_memory)*100, 1)
            else:
                gpu_use_memory = None
                gpu_use_percent = None

            gpu_instance['total_size'] = gpu_total_memory
            gpu_instance['use_size'] = gpu_use_memory
            gpu_instance['use_percent'] = gpu_use_percent

        except Exception as ex:
            print(f"[SystemInfo] gpu_info error : {str(ex)}")

        return gpu_instance

    def memory_info(self):
        '''
        메모리 정보 조회
        '''
        try:
            ps_cmd = 'Get-CimInstance -ClassName Win32_PhysicalMemory | Select-Object -Property Manufacturer, PartNumber, Speed, Capacity'
            memory_info = PowerShell.convert_to_json(cmd=ps_cmd)

            memory_capacity = 0
            memory_instance = {'name': []}
            if isinstance(memory_info, dict):
                memory_info['Capacity'] //= 1048576  # MB
                memory_name = f"{memory_info['Manufacturer']} {memory_info['PartNumber'].strip()} {memory_info['Speed']}MHz {memory_info['Capacity']//1024}GB"
                memory_instance['name'].append(memory_name)

                memory_capacity += memory_info['Capacity']
            else:
                for info in memory_info:
                    info['Capacity'] //= 1048576  # MB
                    memory_name = f"{info['Manufacturer']} {info['PartNumber'].strip()} {info['Speed']}MHz {info['Capacity']//1024}GB"
                    memory_instance['name'].append(memory_name)

                    memory_capacity += info['Capacity']

            memory_instance['total_size'] = memory_capacity

        except Exception as ex:
            print(f"[SystemInfo] memory_info error : {str(ex)}")

        return memory_instance

    def disk_info(self):
        '''
        디스크 정보 조회
        '''
        try:
            ps_cmd = 'Get-CimInstance -ClassName Win32_DiskDrive | Select-Object -Property Index, Model, Size'
            disk_info = PowerShell.convert_to_json(cmd=ps_cmd)

            disk_instance = {'name': []}
            if isinstance(disk_info, dict):
                disk_info['Size'] //= 1048576  # MB
                disk_name = f"{disk_info['Model']} {disk_info['Size']//1024}GB"
                disk_instance['name'].append(disk_name)
            else:
                for info in disk_info:
                    info['Size'] //= 1048576
                    disk_name = f"{info['Model']} {info['Size']//1024}GB"
                    disk_instance['name'].append(disk_name)

        except Exception as ex:
            print(f"[SystemInfo] disk_info error : {str(ex)}")

        return disk_instance

    def volume_info(self):
        '''
        볼륨 정보 조회
        '''
        try:
            ps_cmd = "Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' | Select-Object -Property Name, FileSystem, Size, FreeSpace"
            # DriveType 3 (WMI에서 고정 하드 디스크에 사용하는 값)
            volume_info = PowerShell.convert_to_json(cmd=ps_cmd)

            volume_instance = []
            if isinstance(volume_info, dict):
                volume_info['Size'] //= 1048576  # MB
                volume_info['FreeSpace'] //= 1048576  # MB

                volume_use_size = volume_info['Size'] - volume_info['FreeSpace']
                volume_use_percent = round((volume_use_size / volume_info['Size'])*100, 1)
                volume = {
                    'name': volume_info['Name'][0],
                    'file_system': volume_info['FileSystem'],
                    'total_size': volume_info['Size'],
                    'free_space': volume_info['FreeSpace'],
                    'use_percent': volume_use_percent,
                }
                volume_instance.append(volume)
            else:
                for info in volume_info:
                    info['Size'] //= 1048576  # MB
                    info['FreeSpace'] //= 1048576  # MB

                    volume_use_size = info['Size'] - info['FreeSpace']
                    volume_use_percent = round((volume_use_size / info['Size'])*100, 1)
                    volume = {
                        'name': info['Name'][0],
                        'file_system': info['FileSystem'],
                        'total_size': info['Size'],
                        'free_space': info['FreeSpace'],
                        'use_percent': volume_use_percent,
                    }
                    volume_instance.append(volume)

        except Exception as ex:
            print(f"[SystemInfo] volume_info error : {str(ex)}")

        return volume_instance


class SystemMonitoring:
    '''
    시스템 사용량 모니터링
    '''
    def __init__(self):
        pass

    def run(self):
        try:
            with open('./system_info.json', 'r') as json_file:
                info = json.loads(json_file.read())

            info['monitoring'] = 'usage'
            self.cpu_usage(info['cpu'])
            self.gpu_usage(info['gpu'])
            self.memory_usage(info['memory'])
            self.volume_usage(info['volume'])

            with open('./system_info.json', 'w') as json_file:
                json_file.write(json.dumps(info))

        except Exception as ex:
            print(f"[SystemMonitoring] run error : {str(ex)}")

        return info

    def cpu_usage(self, cpu_instance):
        '''
        CPU 사용률
        '''
        try:
            ps_cmd = "(Get-CimInstance -ClassName Win32_Processor).LoadPercentage"
            cpu_use_percent = PowerShell.isnumeric(cmd=ps_cmd)

            cpu_instance['use_percent'] = cpu_use_percent

        except Exception as ex:
            print(f"[SystemMonitoring] cpu_info error : {str(ex)}")

    def gpu_usage(self, gpu_instance):
        '''
        GPU 사용량
        '''
        try:
            ps_cmd = "(((Get-Counter '\GPU Process Memory(*)\Local Usage').CounterSamples | where CookedValue).CookedValue | measure -sum).sum"
            gpu_use_memory = PowerShell.isnumeric(cmd=ps_cmd)

            if isinstance(gpu_use_memory, int):
                gpu_use_memory //= 1048576  # MB
                gpu_use_percent = round((gpu_use_memory / gpu_instance['total_size']) * 100, 1)
            else:
                gpu_use_memory = None
                gpu_use_percent = None

            gpu_instance['use_size'] = gpu_use_memory
            gpu_instance['use_percent'] = gpu_use_percent

        except Exception as ex:
            print(f"[SystemMonitoring] gpu_info error : {str(ex)}")

    def memory_usage(self, memory_instance):
        '''
        메모리 사용량
        '''
        try:
            ps_cmd = "(Get-CimInstance -Class Win32_OperatingSystem).FreePhysicalMemory"
            memory_free_size = PowerShell.isnumeric(cmd=ps_cmd)

            if isinstance(memory_free_size, int):
                memory_free_size //= 1024  # MB
                memory_use_size = memory_instance['total_size'] - memory_free_size
                memory_use_percent = round((memory_use_size / memory_instance['total_size']) * 100, 1)
            else:
                memory_free_size = None
                memory_use_size = None
                memory_use_percent = None
            
            memory_instance['use_size'] = memory_use_size
            memory_instance['use_percent'] = memory_use_percent

        except Exception as ex:
            print(f"[SystemMonitoring] memory_info error : {str(ex)}")

    def volume_usage(self, volume_instance):
        '''
        볼륨 사용량
        '''
        try:
            ps_cmd = "Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' | Select-Object -Property Name, FreeSpace"
            volume_info = PowerShell.convert_to_json(cmd=ps_cmd)

            for i in range(len(volume_instance)):
                volume_info[i]['FreeSpace'] //= 1048576  # MB
                volume_use_size = volume_instance[i]['total_size'] - volume_info[i]['FreeSpace']
                volume_use_percent = round((volume_use_size / volume_instance[i]['total_size']) * 100, 1)

                volume_instance[i]['free_space'] = volume_info[i]['FreeSpace']
                volume_instance[i]['use_percent'] = volume_use_percent

        except Exception as ex:
            print(f"[SystemMonitoring] volume_info error : {str(ex)}")
