﻿// https://www.codeproject.com/Articles/528178/Load-DLL-From-Embedded-Resource

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;

public class DLLService
{
    private static Dictionary<string, Assembly> dic = null;

    /// <summary>
    /// Accepts a string array with strings in the format of
    /// "ProgramName.Folders.NameOfDll.dll"
    /// i.e. "InstallerUpdater.Assemblies.InstallerUpdater.dll"
    /// </summary>
    /// <param name="resourcesToLoad"></param>
    public static void LoadDLLsFromEmbeddedResources(List<string> resourcesToLoad)
    {
        foreach (var resourceToLoad in resourcesToLoad)
        {
            Load(resourceToLoad, resourceToLoad);
        }
        AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(FetchAssembliesFromMemory);
    }

    private static Assembly FetchAssembliesFromMemory(object sender, ResolveEventArgs args)
    {
        return Get(args.Name);
    }

    public static void Load(string embeddedResource, string fileName)
    {
        if (dic == null)
            dic = new Dictionary<string, Assembly>();

        byte[] ba = null;
        Assembly asm = null;
        Assembly curAsm = Assembly.GetExecutingAssembly();

        using (Stream stm = curAsm.GetManifestResourceStream(embeddedResource))
        {
            if (stm == null)
            {
                throw new Exception(embeddedResource + " is not found in Embedded Resources.");
            }

            ba = new byte[(int)stm.Length];
            stm.Read(ba, 0, (int)stm.Length);

            try
            {
                asm = Assembly.Load(ba);
                dic.Add(asm.FullName, asm);
                return;
            }
            catch
            {
            }
        }

        bool fileOk = false;
        string tempFile = "";

        using (SHA1CryptoServiceProvider sha1 = new SHA1CryptoServiceProvider())
        {
            string fileHash = BitConverter.ToString(sha1.ComputeHash(ba)).Replace("-", string.Empty); ;

            tempFile = Path.GetTempPath() + fileName;

            if (File.Exists(tempFile))
            {
                byte[] bb = File.ReadAllBytes(tempFile);
                string fileHash2 = BitConverter.ToString(sha1.ComputeHash(bb)).Replace("-", string.Empty);

                if (fileHash == fileHash2)
                {
                    fileOk = true;
                }
                else
                {
                    fileOk = false;
                }
            }
            else
            {
                fileOk = false;
            }
        }

        if (!fileOk)
        {
            File.WriteAllBytes(tempFile, ba);
        }

        asm = Assembly.LoadFile(tempFile);

        dic.Add(asm.FullName, asm);
    }

    public static Assembly Get(string assemblyFullName)
    {
        if (dic == null || dic.Count == 0)
            return null;

        if (dic.ContainsKey(assemblyFullName))
            return dic[assemblyFullName];

        return null;
    }
}
