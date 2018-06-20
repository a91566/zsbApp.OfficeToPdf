using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.IO;

namespace zsbApp.OfficeToPdf.Test
{
    public class RegeditTool
    {

      //  #region 查询注册表,判断本机是否安装office2003,2007和wps
        public string ExistsRegedit()
        {
            // RegistryKey rk = Registry.LocalMachine;
            RegistryKey localKey;
            if (Environment.Is64BitOperatingSystem)
                localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            else
                localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);

            RegistryKey akey2016 = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot\");//查询2016
            RegistryKey akey2013 = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\Excel\InstallRoot\");//查询2013
            RegistryKey akey2010 = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\Excel\InstallRoot\");//查询2010
            RegistryKey akey2007 = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\Excel\InstallRoot\");//查询2007
            RegistryKey akey2003 = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\11.0\Excel\InstallRoot\");//查询2003
            RegistryKey sss = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\Common\FileIO\OfficeCollaborationSyncIntegrationEnabled");//查询2003
          
            if (akey2016 != null && akey2016.GetValue("Path") != null)
            {
                return "Office2016";
            }
            else if (akey2013 != null && akey2013.GetValue("Path") != null)
            {
                return "Office2013";
            }
            else if (akey2010 != null && akey2010.GetValue("Path") != null)
            {
                return "Office2010";
            }
            else if (akey2007 != null && akey2007.GetValue("Path") != null)
            {
                return "Office2007";
            }
            else if (akey2003!= null && akey2003.GetValue("Path") != null)
            {
                return "Office2003";
            }
            else
            {
                Type type2 = Type.GetTypeFromProgID("Ket.Application");//V9版本类型
                if (type2 != null)
                {
                    return "WPS2013";
                }

                Type type3 = Type.GetTypeFromProgID("Kwps.Application");//V10版本类型
                if (type3 != null)
                {
                    return "WPS2016";
                }
            }

            return "";
  

            ////检查本机是否安装Office2010
            //if (akey2013 != null)
            //{
            //    string file10 = akey2013.GetValue("Path").ToString();
            //    if (File.Exists(file10 + "Excel.exe"))
            //    {
            //        return "Office2013";
            //    }
            //}
            ////检查本机是否安装Office2013        
            //if (akey2010 != null)
            //{
            //    string file13 = akey2010.GetValue("Path").ToString();
            //    if (File.Exists(file13 + "Excel.exe"))
            //    {
            //        return "Office2010";
            //    }
            //}


            ////检查本机是否安装wps
            //if (akeytwo != null)
            //{
            //    string filewps = akeytwo.GetValue("InstallRoot").ToString();
            //    if (File.Exists(filewps + @"\office6\et.exe"))
            //    {
            //        ifused += 4;
            //    }
            //}


         
     
        }
    }
}
