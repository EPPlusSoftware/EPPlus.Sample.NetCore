/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EPPlusSamples
{
    public class FileInputUtil
    {

        /// <summary>
        /// Returns a fileinfo with the full path of the requested file
        /// </summary>
        /// <param name="directory">A subdirectory</param>
        /// <param name="file"></param>
        /// <returns></returns>
        public static FileInfo GetFileInfo(string directory, string file)
        {
            var rootDir = GetRootDirectory().FullName;
            return new FileInfo(Path.Combine(rootDir, directory, file));
        }

        public static DirectoryInfo GetRootDirectory()
        {
            var currentDir = AppDomain.CurrentDomain.BaseDirectory;
            while (!currentDir.EndsWith("bin"))
            {
                currentDir = Directory.GetParent(currentDir).FullName.TrimEnd('\\');
            }
            return new DirectoryInfo(currentDir).Parent;
        }

        public static DirectoryInfo GetSubDirectory(string directory, string subDirectory)
        {
            var currentDir = GetRootDirectory().FullName;
            return new DirectoryInfo(Path.Combine(currentDir, directory, subDirectory));
        }
    }
}
