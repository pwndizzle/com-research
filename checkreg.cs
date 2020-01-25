// POC code to enumerate COM registry keys that reference User/AppData/ProgramData paths as well as all HKCU entries
// Adapted from the code here - https://stackoverflow.com/questions/34344975/loop-through-registry-and-get-all-subkeys-and-the-values-in-registrykey-in-c-sha

using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace checkreg
{
    class Program
    {
        static IDictionary<string, string> hkcrDict = new Dictionary<string, string>();
        static IDictionary<string, string> hkcuDict = new Dictionary<string, string>();

        static private void processValueNames(RegistryKey Key)
        { //function to process the valueNames for a given key
            
            string[] valuenames = Key.GetValueNames();
            if (valuenames == null || valuenames.Length <= 0) //has no values
                return;
            foreach (string valuename in valuenames)
            {
                if (Key.Name.Contains("LocalServer32") || Key.Name.Contains("InprocServer32") || Key.Name.Contains("InProcServer32") && valuename == "")
                {
                    object obj = Key.GetValue("");
                    if (obj != null)
                        
                        //Console.WriteLine(Key.Name + " - " + obj.ToString());
                        if (Key.Name.Contains("HKEY_CLASSES_ROOT") && !hkcrDict.ContainsKey(obj.ToString()))
                        {
                            hkcrDict.Add(obj.ToString(), Key.Name);
                        }
                        if (Key.Name.Contains("HKEY_CURRENT_USER") && !hkcuDict.ContainsKey(obj.ToString()))
                        {
                            hkcuDict.Add(obj.ToString(), Key.Name);
                        }
                }
            }
        }

        static public void OutputRegKey(RegistryKey Key)
        {
            try
            {
                string[] subkeynames = Key.GetSubKeyNames(); //means deeper folder
                if (subkeynames == null || subkeynames.Length <= 0)
                { //has no more subkey, process
                    processValueNames(Key);
                    return;
                }
                foreach (string keyname in subkeynames)
                { //has subkeys, go deeper
                    using (RegistryKey key2 = Key.OpenSubKey(keyname))
                        OutputRegKey(key2);
                }
                processValueNames(Key);
            }
            catch (Exception e)
            {
                //error, do something
            }
        }

        static void Main(string[] args)
        {
            // Check for existence of User/Appdata/Programdata entries
            RegistryKey hkcrKey = Registry.ClassesRoot.OpenSubKey(@"CLSID");
            OutputRegKey(hkcrKey);

            Console.WriteLine("#################################");
            Console.WriteLine("Temp Folder CLSIDs");
            Console.WriteLine("#################################");
            foreach (KeyValuePair<string, string> entry in hkcrDict)
            {
                if (entry.Key.ToLower().Contains("\\users\\") || entry.Key.ToLower().Contains("\\appdata\\") || entry.Key.ToLower().Contains("\\programdata\\")) {
                //if (entry.Key.ToLower().Contains("program files")){
                //if (entry.Key.ToLower().Contains("system32")){
                    Console.WriteLine(entry.Value + " - " + entry.Key);
                }                
            }

            // Check for existence of HKCU entries
            RegistryKey hkcuKey = Registry.CurrentUser.OpenSubKey(@"Software\Classes\CLSID");
            OutputRegKey(hkcuKey);
            Console.WriteLine("#################################");
            Console.WriteLine("HKCU CLSIDs");
            Console.WriteLine("#################################");
            foreach (KeyValuePair<string, string> entry in hkcuDict)
            {
                Console.WriteLine(entry.Value + " - " + entry.Key);
            }
        }
    }
}
