using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using Microsoft.Web.Administration;

namespace OMMP.Common
{
    /// <summary>
    /// 升级工具
    /// </summary>
    public static class UpdateHelper
    {
        /// <summary>
        /// 项目注册
        /// </summary>
        /// <returns></returns>
        public static SysReStart Register(string name, string port, string url,string dllName="", string code = "")
        {
            SysReStart data = new SysReStart();
            // 程序池名称
            data.ApplicationName = Process.GetCurrentProcess().MainModule.ModuleName;
            data.ServiceCode = code;
            if (data.ApplicationName== "w3wp.exe")
            {
                data.RunType = RunType.IIS;
                data.ApplicationName = Environment.GetEnvironmentVariable("APP_POOL_ID", EnvironmentVariableTarget.Process);
            }
            else
            {
                data.RunType = RunType.EXE;
                // 进程exe名称
                data.ApplicationName = Process.GetCurrentProcess().MainModule.ModuleName;
            }
            // 程序路径
            data.ApplicationPath = System.AppDomain.CurrentDomain.BaseDirectory;
            // 服务名称
            data.ServiceName = name;
            // 端口号
            data.ListenPort = port;
            //操作系统
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                data.SystemType = SystemType.Linux;
                data.ApplicationName = dllName;
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                data.SystemType = SystemType.Windows;
            }
            // 本地ip
            var addressList = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName()).AddressList;
            data.IpAddress = addressList.FirstOrDefault(address => address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)?.ToString();
            return data;
        }
        // 启动
        const string APP_POOL_CMD_START = "start";
        // 停止
        const string APP_POOL_CMD_STOP = "stop";

        /// <summary>
        /// 启动指定的应用程序池
        /// </summary>
        /// <param name="appPoolName">待启动的应用程序池名称</param>
        /// <returns>启用应用程序池的结果。true：启动成功；false：启动失败</returns>
        public static bool StartIIS(string appPoolName)
        {
            return ExecuteIISCmd(APP_POOL_CMD_START, appPoolName);
        }

        /// <summary>
        /// 停止应用程序池
        /// </summary>
        /// <param name="appPoolName">待停止的应用程序池的名称</param>
        /// <returns>停止应用程序池的结果。true：停止成功；false：停止失败</returns>
        public static bool StopIIS(string appPoolName)
        {
            return ExecuteIISCmd(APP_POOL_CMD_STOP, appPoolName);
        }
        /// <summary>
        /// 执行应用程序池相关的命令
        /// </summary>
        /// <param name="method"></param>
        /// <param name="appPoolName"></param>
        /// <returns></returns>
        private static bool ExecuteIISCmd(string method, string appPoolName)
        {
            try
            {
                var serverMgr = new ServerManager();
                var pool = serverMgr.ApplicationPools[appPoolName];
                if (pool == null)
                {
                    throw new Exception("停止应用程序池失败，没有找到相应的应用程序池");
                }

                if (method == APP_POOL_CMD_STOP)
                {
                    // 当前应用程序池处于停止状态
                    if (pool.State == ObjectState.Stopped)
                        return true;

                    var stat = pool.Stop();

                    var now = DateTime.Now;
                    var span = DateTime.Now - now;
                    while (span.TotalMinutes <= 2)
                    {
                        if (pool.State == ObjectState.Stopped)
                            break;
                        System.Threading.Thread.Sleep(50);
                        span = DateTime.Now - now;
                    }

                    return (pool.State == ObjectState.Stopped);
                }
                else
                {
                    if (pool.State == ObjectState.Started)
                        return true;
                    var stat = pool.Start();
                    var now = DateTime.Now;
                    var span = DateTime.Now - now;
                    while (span.TotalMinutes <= 2)
                    {
                        if (pool.State == ObjectState.Started)
                            break;
                        System.Threading.Thread.Sleep(50);
                        span = DateTime.Now - now;
                    }

                    return (pool.State == ObjectState.Started);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }

        /// <summary>
        /// 启动指定的进程
        /// </summary>
        /// <param name="appPoolName">待启动的应用指定的进程名称</param>
        /// <returns>启用应用指定的进程的结果。true：启动成功；false：启动失败</returns>
        public static bool StartWindowsEXE(string appPoolName, string path)
        {
            return ExecuteWindowsEXECmd(APP_POOL_CMD_START, appPoolName,path);
        }

        /// <summary>
        /// 停止指定的进程
        /// </summary>
        /// <param name="appPoolName">待停止的应用指定的进程的名称</param>
        /// <returns>停止应用指定的进程的结果。true：停止成功；false：停止失败</returns>
        public static bool StopWindowsEXE(string appPoolName, string path)
        {
            return ExecuteWindowsEXECmd(APP_POOL_CMD_STOP, appPoolName, path);
        }
        /// <summary>
        /// 执行应用程序池相关的命令
        /// </summary>
        /// <param name="method"></param>
        /// <param name="appPoolName"></param>
        /// <returns></returns>
         private static bool ExecuteWindowsEXECmd(string method, string appPoolName,string path) 
        {
            try
            {
                var exeName = appPoolName.Replace(".exe", "");
                // 先获取当前进程状态
                var runBool = false;
                // 判断是否有服务
                var serviceBool = ISWindowsServiceInstalled(exeName);
                if (serviceBool)
                {
                    // 服务
                    runBool = ISStart(exeName);
                    if (method == APP_POOL_CMD_STOP)
                    {
                        if (runBool == false)
                        {
                            return true;
                        }
                        // 停止
                        StopService(exeName);
                    }
                    else
                    {
                        if (runBool == true)
                        {
                            return true;
                        }
                        // 启动
                        StartService(exeName);
                    }
                }
                else
                {
                    // 进程
                    var processList = Process.GetProcessesByName(exeName).ToList();
                    runBool = processList.Count > 0;
                    if (method == APP_POOL_CMD_STOP)
                    {
                        if (runBool == false)
                        {
                            return true;
                        }
                        foreach (var item in processList)
                        {
                            // 关闭
                            item.Kill();
                        }
                    }
                    else
                    {
                        if (runBool == true)
                        {
                            return true;
                        }
                        // 启动
                        var fullName = path + appPoolName;
                        ProcessStartInfo psi = new ProcessStartInfo();
                        psi.FileName = fullName;
                        psi.UseShellExecute = true;
                        psi.WorkingDirectory = path;
                        psi.CreateNoWindow = false;
                        Process.Start(psi);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }
        /// <summary>
        /// 判断是否存在服务
        /// </summary>
        /// <param name="serviceName"></param>
        public static bool ISWindowsServiceInstalled(string serviceName)
        {
            // get list of Windows services
            ServiceController[] services = ServiceController.GetServices();
            foreach (ServiceController service in services)
            {
                if (service.ServiceName == serviceName)
                    return true;
            }
            return false;
        }
        /// <summary>
        /// 启动某个服务
        /// </summary>
        /// <param name="serviceName"></param>
        public static void StartService(string serviceName)
        {
            try
            {
                ServiceController[] services = ServiceController.GetServices();
                foreach (ServiceController service in services)
                {
                    if (service.ServiceName == serviceName)
                    {
                        service.Start();
                        service.WaitForStatus(ServiceControllerStatus.Running, new TimeSpan(0, 0, 30));
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 停止某个服务
        /// </summary>
        /// <param name="serviceName"></param>
        public static void StopService(string serviceName)
        {
            try
            {
                ServiceController[] services = ServiceController.GetServices();
                foreach (ServiceController service in services)
                {
                    if (service.ServiceName == serviceName)
                    {
                        service.Stop();


                        service.WaitForStatus(ServiceControllerStatus.Running, new TimeSpan(0, 0, 30));
                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// 判断某个服务是否启动
        /// </summary>
        /// <param name="serviceName"></param>
        public static bool ISStart(string serviceName)
        {
            bool result = true;
            try
            {
                ServiceController[] services = ServiceController.GetServices();
                foreach (ServiceController service in services)
                {
                    if (service.ServiceName == serviceName)
                    {
                        if ((service.Status == ServiceControllerStatus.Stopped)
                            || (service.Status == ServiceControllerStatus.StopPending))
                        {
                            result = false;
                        }
                    }
                }
            }
            catch { }
            return result;
        }

        /// <summary>
        /// 启动指定的进程
        /// </summary>
        /// <param name="appPoolName">待启动的应用指定的进程名称</param>
        /// <returns>启用应用指定的进程的结果。true：启动成功；false：启动失败</returns>
        public static bool StartLinuxEXE(string appPoolName, string applicationPath, string listenPort)
        {
            return ExecuteLinuxEXECmd(APP_POOL_CMD_START, appPoolName, applicationPath, listenPort);
        }

        /// <summary>
        /// 停止指定的进程
        /// </summary>
        /// <param name="appPoolName">待停止的应用指定的进程的名称</param>
        /// <returns>停止应用指定的进程的结果。true：停止成功；false：停止失败</returns>
        public static bool StopLinuxEXE(string appPoolName, string applicationPath, string listenPort)
        {
            return ExecuteLinuxEXECmd(APP_POOL_CMD_STOP, appPoolName,applicationPath,listenPort);
        }
        /// <summary>
        /// 执行应用程序池相关的命令
        /// </summary>
        /// <param name="method"></param>
        /// <param name="appPoolName"></param>
        /// <returns></returns>
        private static bool ExecuteLinuxEXECmd(string method, string appPoolName, string applicationPath, string listenPort)
        {
            try
            {
                // 先获取当前进程状态
                var runBool = false;
                // 服务
                runBool = ISLinux(appPoolName, applicationPath, listenPort);
                if (method == APP_POOL_CMD_STOP)
                {
                    if (runBool == false)
                    {
                        return true;
                    }
                    // 停止
                    StopLinux( appPoolName, applicationPath, listenPort);
                }
                else
                {
                    if (runBool == true)
                    {
                        return true;
                    }
                    // 启动
                    StartLinux( appPoolName, applicationPath, listenPort);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }
        /// <summary>
        /// 启动某个服务
        /// </summary>
        /// <param name="serviceName"></param>
        public static void StartLinux(string appPoolName, string applicationPath, string listenPort)
        {
            try
            {
                //合并文件命令
                //string command = "nohup dotnet /opt/ommp/masterData/OMMP.MasterData.Web.dll > /opt/ommp/masterData/log/masterData.log 2>&1 &";
                string command = "nohup dotnet "+ applicationPath + appPoolName + " > " + applicationPath + "nohup_log.log 2>&1 &";
                //执行结果
                string result = "";
                using (System.Diagnostics.Process proc = new System.Diagnostics.Process())
                {
                    string command1 = "cd "+ applicationPath;
                    proc.StartInfo.FileName = "/bin/bash";
                    proc.StartInfo.Arguments = "-c \"" + command1 + ";" + command + " \"";
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.Start();
                    result += proc.StandardOutput.ReadToEnd();
                    result += proc.StandardError.ReadToEnd();
                    Console.WriteLine(result);
                    proc.WaitForExit();
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw new Exception(ex.Message.ToString());
            }
        }

        /// <summary>
        /// 停止某个服务
        /// </summary>
        /// <param name="serviceName"></param>
        public static void StopLinux(string appPoolName, string applicationPath, string listenPort)
        {
            try
            {
                //拼接合并命令中的文件字符串,sourcePath为文件块所在目录,targetPath为合并文件的目录
                string command = "kill `netstat -nlp | grep :"+listenPort+" | awk '{print $7}' | awk -F\"/\" '{ print $1 }'`";
                //执行结果
                string result = "";
                using (System.Diagnostics.Process proc = new System.Diagnostics.Process())
                {
                    proc.StartInfo.FileName = "/bin/bash";
                    proc.StartInfo.Arguments = "-c \"" + command + " \"";
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.Start();
                    result += proc.StandardOutput.ReadToEnd();
                    result += proc.StandardError.ReadToEnd();
                    Console.WriteLine(result);
                    proc.WaitForExit();
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw new Exception(ex.Message.ToString());
            }
        }
        /// <summary>
        /// 判断某个服务是否启动
        /// </summary>
        /// <param name="serviceName"></param>
        public static bool ISLinux(string appPoolName, string applicationPath, string listenPort)
        {
            bool result = true;
            try
            {
                using (System.Diagnostics.Process proc = new System.Diagnostics.Process())
                {
                    var command1 = "netstat -nlp | grep :" + listenPort;
                    proc.StartInfo.FileName = "/bin/bash";
                    proc.StartInfo.Arguments = "-c \"" + command1 + " \"";
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.Start();

                    var result1 = proc.StandardOutput.ReadToEnd();
                    result1 += proc.StandardError.ReadToEnd();
                    if (string.IsNullOrEmpty(result1))
                    {
                        return false;
                    }
                    proc.WaitForExit();
                }
            }
            catch { }
            return result;
        }
    }
    /// <summary>
    /// 系统重启
    /// </summary>
    public class SysReStart
    {
        /// <summary>
        /// 服务编号
        /// </summary>
        public string ServiceCode { get; set; } = "服务编号";
        /// <summary>
        /// 服务名称
        /// </summary>
        public string ServiceName { get; set; } = "服务名称";
        /// <summary>
        /// 应用程序名
        /// </summary>
        public string ApplicationName { get; set; } = "应用程序名";
        /// <summary>
        /// 程序路径
        /// </summary>
        public string ApplicationPath { get; set; } = "程序路径";
        /// <summary>
        /// 端口号
        /// </summary>
        public string ListenPort { get; set; } = "端口号";
        /// <summary>
        /// 更新时间
        /// </summary>
        public DateTime UpdateTime { get; set; } = DateTimeHelper.Minimum;
        /// <summary>
        /// 启动时间
        /// </summary>
        public DateTime StartTime { get; set; } = DateTime.Now;
        /// <summary>
        /// 更新版本号
        /// </summary>
        public string Version { get; set; } = "";
        /// <summary>
        /// 本地ip
        /// </summary>
        public string IpAddress { get; set; } = "本地ip";
        /// <summary>
        /// 本地系统
        /// </summary>
        public SystemType SystemType { get; set; } = SystemType.Windows;
        /// <summary>
        /// 运行类型
        /// </summary>
        public RunType RunType { get; set; } = RunType.IIS;
        /// <summary>
        /// 0后，1前
        /// </summary>
        public int QianOrHou { get; set; } = 0;
    }
    /// <summary>
    /// 运行类型
    /// </summary>
    public enum RunType
    {
        /// <summary>
        /// iis
        /// </summary>
        IIS,
        /// <summary>
        /// 本地exe
        /// </summary>
        EXE,
    }
    /// <summary>
    /// 操作系统
    /// </summary>
    public enum SystemType
    {
        /// <summary>
        /// Windows
        /// </summary>
        Windows,
        /// <summary>
        /// Linux
        /// </summary>
        Linux,
    }
}
