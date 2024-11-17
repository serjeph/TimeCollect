using System;
using System.IO;
using System.Security.AccessControl;

namespace TimeCollect.Helpers
{
    public class Security
    {
        public static void CreateDirectoryAccess(string folderPath)
        {
            string userName = Environment.UserName;

            // Get a DirectorySecurity object that represents the
            // current security settings.
            DirectorySecurity directorySecurity = Directory.GetAccessControl(folderPath);

            FileSystemAccessRule accessRule = new FileSystemAccessRule(
                    userName,
                    FileSystemRights.CreateDirectories,
                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags.None,
                    AccessControlType.Allow
                );

            directorySecurity.AddAccessRule(accessRule);

            // Set the new access control list on the directory.
            Directory.SetAccessControl(folderPath, directorySecurity);
        }
    }
}
