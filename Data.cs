using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ExcelCompareTool
{
    [DataContract]
    public class Data
    {
        [DataMember]
        public string ConnectionString { get; set; }
        public void Serializer()
        {

            try
            {
                DataContractJsonSerializer formatter = new DataContractJsonSerializer(typeof(Data));
                using (MemoryStream stream = new MemoryStream())
                {
                    formatter.WriteObject(stream, this);
                    string result = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                    File.WriteAllText(GetPath(), result);
                }
            }
            catch (System.Exception e)
            {
                MessageBox.Show("持久化配置失败！");
            }
        }
        public static Data Get()
        {
            try
            {
                var str = File.ReadAllText(GetPath());

                DataContractJsonSerializer formatter = new DataContractJsonSerializer(typeof(Data));
                using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(str)))
                {
                    return formatter.ReadObject(stream) as Data;
                }
            }
            catch (System.Exception e)
            {
            }
            return null;
        }
        public static string GetPath()
        {
            RegistryKey folders = OpenRegistryPath(Registry.CurrentUser, @"\software\microsoft\windows\currentversion\explorer\shell folders");
            string pathemp = folders.GetValue("Desktop").ToString();
            if (Directory.Exists(pathemp))
            {
                DirectoryInfo dir = new DirectoryInfo(pathemp);
                string pathnew = dir.Parent.FullName + @"\ERP\User\";
                if (!Directory.Exists(pathnew))
                {
                    Directory.CreateDirectory(pathnew);
                }
                pathnew = pathnew + "XKExcelImport.dat";
                return pathnew;
            }
            else
            {
                string pathnew = @"C:\XK\";
                if (!Directory.Exists(pathnew))
                {
                    Directory.CreateDirectory(pathnew);
                }
                return @"C:\XK\XKExcelCompare.dat";
            }

        }

        private static RegistryKey OpenRegistryPath(RegistryKey root, string s)
        {
            s = s.Remove(0, 1) + @"/";
            while (s.IndexOf(@"/") != -1)
            {
                root = root.OpenSubKey(s.Substring(0, s.IndexOf(@"/")));
                s = s.Remove(0, s.IndexOf(@"/") + 1);
            }
            return root;
        }
    }
}
