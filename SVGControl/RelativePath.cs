using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using PInvoke = Windows.Win32.PInvoke;

#nullable enable

[assembly: InternalsVisibleTo("SVGControl.Test")]
namespace SVGControl
{
    internal static class RelativePath
    {
     
        /// <summary>
        /// Creates a relative path from one file or folder to another.
        /// </summary>
        /// <param name="anchorPath">Contains the directory that defines the start of the relative path.</param>
        /// <param name="pathToMakeRelative">Contains the path that defines the endpoint of the relative path.</param>
        /// <returns>The relative path from the start directory to the end path or <c>pathToMakeRelative</c> if the paths are not related.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="UriFormatException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public static string MakeRelativePath(this String pathToMakeRelative, String anchorPath)
        {
            if (String.IsNullOrEmpty(anchorPath)) throw new ArgumentNullException("anchorPath");
            if (String.IsNullOrEmpty(pathToMakeRelative)) throw new ArgumentNullException("pathToMakeRelative");

            Uri fromUri = new Uri(anchorPath);
            Uri toUri = new Uri(pathToMakeRelative);

            if (fromUri.Scheme != toUri.Scheme) { return pathToMakeRelative; } // path can't be made relative.

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            String relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            if (toUri.Scheme.Equals("file", StringComparison.InvariantCultureIgnoreCase))
            {
                relativePath = relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            }

            return relativePath;
        }

        public static string GetRelativeURI(this String pathToMakeRelative, String anchorPath)
        {
            if (String.IsNullOrEmpty(anchorPath)) throw new ArgumentNullException("anchorPath");
            if (String.IsNullOrEmpty(pathToMakeRelative)) throw new ArgumentNullException("pathToMakeRelative");

            Uri fromUri = new Uri(anchorPath);
            Uri toUri = new Uri(pathToMakeRelative);

            if (fromUri.Scheme != toUri.Scheme) { return pathToMakeRelative; } // path can't be made relative.

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            String relativePath = Uri.UnescapeDataString(relativeUri.ToString());
            if (!relativePath[0].Equals('.'))
                relativePath = "./" + relativePath;

            return relativePath;
        }

        public static string AbsoluteFromPath(this String pathToMakeAbsolute, String anchorPath)
        {
            if (String.IsNullOrEmpty(anchorPath)) throw new ArgumentNullException("anchorPath");
            if (String.IsNullOrEmpty(pathToMakeAbsolute)) throw new ArgumentNullException("pathToMakeAbsolute");

            anchorPath = NormalizeFolderpath(anchorPath);
            string absolutePath = Path.GetFullPath(anchorPath + pathToMakeAbsolute);

            return absolutePath;
        }

        public static string AbsoluteFromURI(this string uriToMakeAbsolute, string anchorPath)
        {
            if (uriToMakeAbsolute.StartsWith("./") || uriToMakeAbsolute.StartsWith("../"))
            {
                anchorPath = NormalizeFolderpath(anchorPath);
                string relativePath = uriToMakeAbsolute;
                //relativePath = relativePath.Substring(1);
                //relativePath = uriToMakeAbsolute.Substring(2);
                //relativePath = relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

                string absolutePath = GetFullPath(relativePath, anchorPath);
                return absolutePath;
            }
            return uriToMakeAbsolute;
        }

        #region .NET 7.0 GetFullPath decompiled and retrofitted Function with helpers
        /// <summary>
        /// Decompiled from .NET 7.0 and retrofitted
        /// </summary>
        /// <param name="path"></param>
        /// <param name="basePath"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        static public string GetFullPath(string path, string basePath)
        {
            if (string.IsNullOrEmpty(path)) throw new ArgumentNullException("path");
            if (string.IsNullOrEmpty(basePath)) throw new ArgumentNullException("basePath");
            
            if (!IsPathFullyQualified(basePath))
                throw new ArgumentException("Arg_BasePathNotFullyQualified", nameof(basePath));

            if (basePath.Contains('\0') || path.Contains('\0'))
                throw new ArgumentException("Argument_InvalidPathChars");

            if (IsPathFullyQualified(path))
                return Path.GetFullPath(path);

            int length = path.Length;
            string combinedPath;
            if (length >= 1 && IsDirectorySeparator(path[0]))
            {
                // Path is current drive rooted i.e. starts with \:
                // "\Foo" and "C:\Bar" => "C:\Foo"
                // "\Foo" and "\\?\C:\Bar" => "\\?\C:\Foo"
                //combinedPath = Join(Path.GetPathRoot(basePath), path.AsSpan(1)); replaced with Concat2 
                combinedPath = Path.GetPathRoot(basePath).Concat2(path.AsSpan(1)); // Cut the separator to ensure we don't end up with two separators when joining with the root.

            }
            else if (length >= 2 && IsValidDriveChar(path[0]) && path[1] == Path.VolumeSeparatorChar)
            {
                // Drive relative paths
                Debug.Assert(length == 2 || !IsDirectorySeparator(path[2]));

                if (GetVolumeName(path.AsSpan()).Equals(GetVolumeName(basePath.AsSpan()),StringComparison.Ordinal))
                {
                    // Matching root
                    // "C:Foo" and "C:\Bar" => "C:\Bar\Foo"
                    // "C:Foo" and "\\?\C:\Bar" => "\\?\C:\Bar\Foo"
                    combinedPath = new string(basePath.AsSpan().Concat2(path.AsSpan(2)).ToArray());
                }
                else
                {
                    // No matching root, root to specified drive
                    // "D:Foo" and "C:\Bar" => "D:Foo"
                    // "D:Foo" and "\\?\C:\Bar" => "\\?\D:\Foo"
                    combinedPath = !IsDevice(basePath)
                        ? path.Insert(2, @"\")
                        : length == 2
                            ? new string(basePath.AsSpan(0, 4).Concat2(path.AsSpan()).Concat2(@"\".AsSpan()).ToArray())
                            : new string(basePath.AsSpan(0, 4).Concat2(path.AsSpan(0, 2)).Concat2(@"\".AsSpan()).Concat2(path.AsSpan(2)).ToArray());
                }
            }
            else
            {
                // "Simple" relative path
                // "Foo" and "C:\Bar" => "C:\Bar\Foo"
                // "Foo" and "\\?\C:\Bar" => "\\?\C:\Bar\Foo"
                combinedPath = new string(basePath.AsSpan().Concat2(path.AsSpan()).ToArray());
            }

            // Device paths are normalized by definition, so passing something of this format (i.e. \\?\C:\.\tmp, \\.\C:\foo)
            // to Windows APIs won't do anything by design. Additionally, GetFullPathName() in Windows doesn't root
            // them properly. As such we need to manually remove segments and not use GetFullPath().
            return IsDevice(combinedPath)
                ? RemoveRelativeSegments(combinedPath, GetRootLength(combinedPath))
                : GetFullPathInternal(combinedPath);
        }

        // Gets the full path without argument validation
        private static string GetFullPathInternal(string path)
        {
            Debug.Assert(!string.IsNullOrEmpty(path));
            Debug.Assert(!path.Contains('\0'));

            if (IsExtended(path))
            {
                // \\?\ paths are considered normalized by definition. Windows doesn't normalize \\?\
                // paths and neither should we. Even if we wanted to GetFullPathName does not work
                // properly with device paths. If one wants to pass a \\?\ path through normalization
                // one can chop off the prefix, pass it to GetFullPath and add it again.
                return path;
            }

            return Normalize(path);
        }

        internal const int MaxShortPath = 260;
        internal const char VolumeSeparatorChar = ':';

        /// <summary>
        /// Returns true if the path specified is relative to the current drive or working directory.
        /// Returns false if the path is fixed to a specific drive or UNC path.  This method does no
        /// validation of the path (URIs will be returned as relative as a result).
        /// </summary>
        /// <remarks>
        /// Handles paths that use the alternate directory separator.  It is a frequent mistake to
        /// assume that rooted paths (Path.IsPathRooted) are not relative.  This isn't the case.
        /// "C:a" is drive relative- meaning that it will be resolved against the current directory
        /// for C: (rooted, but relative). "C:\a" is rooted and not relative (the current directory
        /// will not be used to modify the path).
        /// </remarks>
        internal static bool IsPartiallyQualified(ReadOnlySpan<char> path)
        {
            if (path.Length < 2)
            {
                // It isn't fixed, it must be relative.  There is no way to specify a fixed
                // path with one character (or less).
                return true;
            }

            if (IsDirectorySeparator(path[0]))
            {
                // There is no valid way to specify a relative path with two initial slashes or
                // \? as ? isn't valid for drive relative paths and \??\ is equivalent to \\?\
                return !(path[1] == '?' || IsDirectorySeparator(path[1]));
            }

            // The only way to specify a fixed path that doesn't begin with two slashes
            // is the drive, colon, slash format- i.e. C:\
            return !((path.Length >= 3)
                && (path[1] == VolumeSeparatorChar)
                && IsDirectorySeparator(path[2])
                // To match old behavior we'll check the drive character for validity as the path is technically
                // not qualified if you don't have a valid drive. "=:\" is the "=" file's default data stream.
                && IsValidDriveChar(path[0]));
        }

        

        /// <summary>
        /// Get the last platform invoke error on the current thread
        /// </summary>
        /// <returns>The last platform invoke error</returns>
        /// <remarks>
        /// The last platform invoke error corresponds to the error set by either the most recent platform
        /// invoke that was configured to set the last error or a call to <see cref="SetLastPInvokeError(int)" />.
        /// </remarks>
        [MethodImpl(MethodImplOptions.InternalCall)]
        public static extern int GetLastPInvokeError();

        internal const int ERROR_SUCCESS = 0x0;
        internal const int ERROR_INVALID_FUNCTION = 0x1;
        internal const int ERROR_FILE_NOT_FOUND = 0x2;
        internal const int ERROR_PATH_NOT_FOUND = 0x3;
        internal const int ERROR_ACCESS_DENIED = 0x5;
        internal const int ERROR_INVALID_HANDLE = 0x6;
        internal const int ERROR_NOT_ENOUGH_MEMORY = 0x8;
        internal const int ERROR_INVALID_DATA = 0xD;
        internal const int ERROR_INVALID_DRIVE = 0xF;
        internal const int ERROR_NO_MORE_FILES = 0x12;
        internal const int ERROR_NOT_READY = 0x15;
        internal const int ERROR_BAD_COMMAND = 0x16;
        internal const int ERROR_BAD_LENGTH = 0x18;
        internal const int ERROR_SHARING_VIOLATION = 0x20;
        internal const int ERROR_LOCK_VIOLATION = 0x21;
        internal const int ERROR_HANDLE_EOF = 0x26;
        internal const int ERROR_NOT_SUPPORTED = 0x32;
        internal const int ERROR_BAD_NETPATH = 0x35;
        internal const int ERROR_NETWORK_ACCESS_DENIED = 0x41;
        internal const int ERROR_BAD_NET_NAME = 0x43;
        internal const int ERROR_FILE_EXISTS = 0x50;
        internal const int ERROR_INVALID_PARAMETER = 0x57;
        internal const int ERROR_BROKEN_PIPE = 0x6D;
        internal const int ERROR_DISK_FULL = 0x70;
        internal const int ERROR_SEM_TIMEOUT = 0x79;
        internal const int ERROR_CALL_NOT_IMPLEMENTED = 0x78;
        internal const int ERROR_INSUFFICIENT_BUFFER = 0x7A;
        internal const int ERROR_INVALID_NAME = 0x7B;
        internal const int ERROR_MOD_NOT_FOUND = 0x7E;
        internal const int ERROR_NEGATIVE_SEEK = 0x83;
        internal const int ERROR_DIR_NOT_EMPTY = 0x91;
        internal const int ERROR_BAD_PATHNAME = 0xA1;
        internal const int ERROR_LOCK_FAILED = 0xA7;
        internal const int ERROR_BUSY = 0xAA;
        internal const int ERROR_ALREADY_EXISTS = 0xB7;
        internal const int ERROR_BAD_EXE_FORMAT = 0xC1;
        internal const int ERROR_ENVVAR_NOT_FOUND = 0xCB;
        internal const int ERROR_FILENAME_EXCED_RANGE = 0xCE;
        internal const int ERROR_EXE_MACHINE_TYPE_MISMATCH = 0xD8;
        internal const int ERROR_FILE_TOO_LARGE = 0xDF;
        internal const int ERROR_PIPE_BUSY = 0xE7;
        internal const int ERROR_NO_DATA = 0xE8;
        internal const int ERROR_PIPE_NOT_CONNECTED = 0xE9;
        internal const int ERROR_MORE_DATA = 0xEA;
        internal const int ERROR_NO_MORE_ITEMS = 0x103;
        internal const int ERROR_DIRECTORY = 0x10B;
        internal const int ERROR_NOT_OWNER = 0x120;
        internal const int ERROR_TOO_MANY_POSTS = 0x12A;
        internal const int ERROR_PARTIAL_COPY = 0x12B;
        internal const int ERROR_ARITHMETIC_OVERFLOW = 0x216;
        internal const int ERROR_PIPE_CONNECTED = 0x217;
        internal const int ERROR_PIPE_LISTENING = 0x218;
        internal const int ERROR_MUTANT_LIMIT_EXCEEDED = 0x24B;
        internal const int ERROR_OPERATION_ABORTED = 0x3E3;
        internal const int ERROR_IO_INCOMPLETE = 0x3E4;
        internal const int ERROR_IO_PENDING = 0x3E5;
        internal const int ERROR_NO_TOKEN = 0x3f0;
        internal const int ERROR_SERVICE_DOES_NOT_EXIST = 0x424;
        internal const int ERROR_EXCEPTION_IN_SERVICE = 0x428;
        internal const int ERROR_PROCESS_ABORTED = 0x42B;
        internal const int ERROR_NO_UNICODE_TRANSLATION = 0x459;
        internal const int ERROR_DLL_INIT_FAILED = 0x45A;
        internal const int ERROR_COUNTER_TIMEOUT = 0x461;
        internal const int ERROR_NO_ASSOCIATION = 0x483;
        internal const int ERROR_DDE_FAIL = 0x484;
        internal const int ERROR_DLL_NOT_FOUND = 0x485;
        internal const int ERROR_NOT_FOUND = 0x490;
        internal const int ERROR_CANCELLED = 0x4C7;
        internal const int ERROR_NETWORK_UNREACHABLE = 0x4CF;
        internal const int ERROR_NON_ACCOUNT_SID = 0x4E9;
        internal const int ERROR_NOT_ALL_ASSIGNED = 0x514;
        internal const int ERROR_UNKNOWN_REVISION = 0x519;
        internal const int ERROR_INVALID_OWNER = 0x51B;
        internal const int ERROR_INVALID_PRIMARY_GROUP = 0x51C;
        internal const int ERROR_NO_SUCH_PRIVILEGE = 0x521;
        internal const int ERROR_PRIVILEGE_NOT_HELD = 0x522;
        internal const int ERROR_INVALID_ACL = 0x538;
        internal const int ERROR_INVALID_SECURITY_DESCR = 0x53A;
        internal const int ERROR_INVALID_SID = 0x539;
        internal const int ERROR_BAD_IMPERSONATION_LEVEL = 0x542;
        internal const int ERROR_CANT_OPEN_ANONYMOUS = 0x543;
        internal const int ERROR_NO_SECURITY_ON_OBJECT = 0x546;
        internal const int ERROR_CANNOT_IMPERSONATE = 0x558;
        internal const int ERROR_CLASS_ALREADY_EXISTS = 0x582;
        internal const int ERROR_NO_SYSTEM_RESOURCES = 0x5AA;
        internal const int ERROR_TIMEOUT = 0x5B4;
        internal const int ERROR_EVENTLOG_FILE_CHANGED = 0x5DF;
        internal const int ERROR_TRUSTED_RELATIONSHIP_FAILURE = 0x6FD;
        internal const int ERROR_RESOURCE_TYPE_NOT_FOUND = 0x715;
        internal const int ERROR_RESOURCE_LANG_NOT_FOUND = 0x717;
        internal const int RPC_S_CALL_CANCELED = 0x71A;
        internal const int ERROR_NOT_A_REPARSE_POINT = 0x1126;
        internal const int ERROR_EVT_QUERY_RESULT_STALE = 0x3AA3;
        internal const int ERROR_EVT_QUERY_RESULT_INVALID_POSITION = 0x3AA4;
        internal const int ERROR_EVT_INVALID_EVENT_DATA = 0x3A9D;
        internal const int ERROR_EVT_PUBLISHER_METADATA_NOT_FOUND = 0x3A9A;
        internal const int ERROR_EVT_CHANNEL_NOT_FOUND = 0x3A9F;
        internal const int ERROR_EVT_MESSAGE_NOT_FOUND = 0x3AB3;
        internal const int ERROR_EVT_MESSAGE_ID_NOT_FOUND = 0x3AB4;
        internal const int ERROR_EVT_PUBLISHER_DISABLED = 0x3ABD;


        /// <summary>
        ///     Returns a string message for the specified Win32 error code.
        /// </summary>
        internal static string GetMessage(int errorCode) =>
            GetMessage(errorCode, IntPtr.Zero);

        private const int FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
        private const int FORMAT_MESSAGE_FROM_HMODULE = 0x00000800;
        private const int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;
        private const int FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x00002000;
        private const int FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;

        [DllImport("KERNEL32.dll", ExactSpelling = true, EntryPoint = "FormatMessageW", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        public static extern unsafe int FormatMessageW(int dwFlags, 
                                                        IntPtr lpSource, 
                                                        uint dwMessageId, 
                                                        int dwLanguageId, 
                                                        void* lpBuffer, 
                                                        int nSize, 
                                                        IntPtr arguments);

        [DllImport("KERNEL32.dll", ExactSpelling = true, EntryPoint = "GetFullPathNameW", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        public static extern unsafe uint GetLongPathNameW(ref char lpszShortPath,
                                                          ref char lpszLongPath,
                                                          uint cchBuffer);

        private static string GetAndTrimString(Span<char> buffer)
        {
            int length = buffer.Length;
            while (length > 0 && buffer[length - 1] <= 32)
            {
                length--; // trim off spaces and non-printable ASCII chars at the end of the resource
            }
            return buffer.Slice(0, length).ToString();
        }

        internal static unsafe string GetMessage(int errorCode, IntPtr moduleHandle)
        {
            int flags = FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ARGUMENT_ARRAY;
            if (moduleHandle != IntPtr.Zero)
            {
                flags |= FORMAT_MESSAGE_FROM_HMODULE;
            }

            // First try to format the message into the stack based buffer.  Most error messages willl fit.
            Span<char> stackBuffer = stackalloc char[256]; // arbitrary stack limit
            fixed (char* bufferPtr = stackBuffer)
            {
                int length = FormatMessageW(flags, moduleHandle, unchecked((uint)errorCode), 0, bufferPtr, stackBuffer.Length, IntPtr.Zero);
                if (length > 0)
                {
                    return GetAndTrimString(stackBuffer.Slice(0, length));
                }
            }

            // We got back an error.  If the error indicated that there wasn't enough room to store
            // the error message, then call FormatMessage again, but this time rather than passing in
            // a buffer, have the method allocate one, which we then need to free.
            if (Marshal.GetLastWin32Error() == ERROR_INSUFFICIENT_BUFFER)
            {
                IntPtr nativeMsgPtr = default;
                try
                {
                    int length = FormatMessageW(flags | FORMAT_MESSAGE_ALLOCATE_BUFFER, moduleHandle, unchecked((uint)errorCode), 0, &nativeMsgPtr, 0, IntPtr.Zero);
                    if (length > 0)
                    {
                        return GetAndTrimString(new Span<char>((char*)nativeMsgPtr, length));
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(nativeMsgPtr);
                }
            }

            // Couldn't get a message, so manufacture one.
            return $"Unknown error (0x{errorCode:x})";
        }

        /// <summary>
        /// Converts the specified Win32 error into a corresponding <see cref="Exception"/> object, optionally
        /// including the specified path in the error message.
        /// </summary>
        internal static Exception GetExceptionForWin32Error(int errorCode, string? path = "")
        {
            // ERROR_SUCCESS gets thrown when another unexpected interop call was made before checking GetLastWin32Error().
            // Errors have to get retrieved as soon as possible after P/Invoking to avoid this.
            Debug.Assert(errorCode != ERROR_SUCCESS);

            switch (errorCode)
            {
                case ERROR_FILE_NOT_FOUND:
                    return new FileNotFoundException(
                        string.IsNullOrEmpty(path) ? SR.IO_FileNotFound : SR.Format(SR.IO_FileNotFound_FileName, path), path);
                case ERROR_PATH_NOT_FOUND:
                    return new DirectoryNotFoundException(
                        string.IsNullOrEmpty(path) ? SR.IO_PathNotFound_NoPathName : SR.Format(SR.IO_PathNotFound_Path, path));
                case ERROR_ACCESS_DENIED:
                    return new UnauthorizedAccessException(
                        string.IsNullOrEmpty(path) ? SR.UnauthorizedAccess_IODenied_NoPathName : SR.Format(SR.UnauthorizedAccess_IODenied_Path, path));
                case ERROR_ALREADY_EXISTS:
                    if (string.IsNullOrEmpty(path))
                        goto default;
                    return new IOException(SR.Format(SR.IO_AlreadyExists_Name, path), MakeHRFromErrorCode(errorCode));
                case ERROR_FILENAME_EXCED_RANGE:
                    return new PathTooLongException(
                        string.IsNullOrEmpty(path) ? SR.IO_PathTooLong : SR.Format(SR.IO_PathTooLong_Path, path));
                case ERROR_SHARING_VIOLATION:
                    return new IOException(
                        string.IsNullOrEmpty(path) ? SR.IO_SharingViolation_NoFileName : SR.Format(SR.IO_SharingViolation_File, path),
                        MakeHRFromErrorCode(errorCode));
                case ERROR_FILE_EXISTS:
                    if (string.IsNullOrEmpty(path))
                        goto default;
                    return new IOException(SR.Format(SR.IO_FileExists_Name, path), MakeHRFromErrorCode(errorCode));
                case ERROR_OPERATION_ABORTED:
                    return new OperationCanceledException();
                case ERROR_INVALID_PARAMETER:
                default:
                    string msg = string.IsNullOrEmpty(path)
                        ? GetPInvokeErrorMessage(errorCode)
                        : $"{GetPInvokeErrorMessage(errorCode)} : '{path}'";
                    return new IOException(
                        msg,
                        MakeHRFromErrorCode(errorCode));
            }

            static string GetPInvokeErrorMessage(int errorCode)
            {
                // Call Kernel32.GetMessage directly in CoreLib. It eliminates one level of indirection and it is necessary to
                // produce correct error messages for CoreCLR Win32 PAL.
#if NET7_0_OR_GREATER && !SYSTEM_PRIVATE_CORELIB
                return Marshal.GetPInvokeErrorMessage(errorCode);
#else
                return GetMessage(errorCode);
#endif
            }
        }

        /// <summary>
        /// If not already an HRESULT, returns an HRESULT for the specified Win32 error code.
        /// </summary>
        internal static int MakeHRFromErrorCode(int errorCode)
        {
            // Don't convert it if it is already an HRESULT
            if ((0xFFFF0000 & errorCode) != 0)
                return errorCode;

            return unchecked(((int)0x80070000) | errorCode);
        }

        /// <summary>
        /// Returns a Win32 error code for the specified HRESULT if it came from FACILITY_WIN32
        /// If not, returns the HRESULT unchanged
        /// </summary>
        internal static int TryMakeWin32ErrorCodeFromHR(int hr)
        {
            if ((0xFFFF0000 & hr) == 0x80070000)
            {
                // Win32 error, Win32Marshal.GetExceptionForWin32Error expects the Win32 format
                hr &= 0x0000FFFF;
            }

            return hr;
        }

        //[DllImport("KERNEL32.dll", ExactSpelling = true, EntryPoint = "GetFullPathNameW", SetLastError = true)]
        //[DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        //public static extern unsafe uint GetFullPathNameW(ref char lpFileName, uint nBufferLength, ref char lpBuffer, [Optional] IntPtr lpFilePart);

        //[DllImport("KERNEL32.dll", ExactSpelling = true, EntryPoint = "GetFullPathNameW", SetLastError = true)]
        //[DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        //static extern uint GetFullPathNameZ(string lpFileName, uint nBufferLength, out StringBuilder lpBuffer, out StringBuilder lpFilePart);

        /// <summary>
        /// Calls GetFullPathName on the given path.
        /// </summary>
        /// <param name="path">The path name. MUST be null terminated after the span.</param>
        /// <param name="builder">Builder that will store the result.</param>
        //private static void GetFullPathName(string path, ref string returnPath)
        //{
        //    // If the string starts with an extended prefix we would need to remove it from the path before we call GetFullPathName as
        //    // it doesn't root extended paths correctly. We don't currently resolve extended paths, so we'll just assert here.
        //    Debug.Assert(IsPartiallyQualified(path.AsSpan()) || !IsExtended(path));

            
            
        //    int bufsz = 1;
        //    StringBuilder sbFull = new StringBuilder(bufsz);          // Full resolved path will go here
        //    StringBuilder sbFile = new StringBuilder(bufsz);          // Filename will go here

        //    uint u = GetFullPathNameZ(path, (uint)bufsz, out sbFull, out sbFile);  // 1st call: Get necessary bufsz
        //    if (u > bufsz)                            // 'u' should be >1
        //    {
        //        bufsz = (int)u + 10;                       // Required _size plus a few
        //        sbFull = new StringBuilder(bufsz);                 // Re-create objects w/ proper _size
        //        sbFile = new StringBuilder(bufsz);                 // "
        //        u = GetFullPathNameZ(path, (uint)bufsz, out sbFull, out sbFile);    // Try again, this should succeed
        //                                                                       // 'sbFull' should now contain "c:\windows\system32\desktop.ini"
        //                                                                       //    and 'sbFile' should contain "desktop.ini"
        //    }
            
            

        //    if (u == 0)
        //    {
        //        // Failure, get the error and throw
        //        int errorCode = GetLastPInvokeError();
        //        if (errorCode == 0)
        //            errorCode = ERROR_BAD_PATHNAME;
        //        throw GetExceptionForWin32Error(errorCode, path.ToString());
        //    }
            
        //    returnPath = sbFull.ToString();

        //}



        ///// <summary>
        ///// Calls GetFullPathName on the given path.
        ///// </summary>
        ///// <param name="path">The path name. MUST be null terminated after the span.</param>
        ///// <param name="builder">Builder that will store the result.</param>
        //private static void GetFullPathName(ReadOnlySpan<char> path, ref ValueStringBuilder builder)
        //{
        //    // If the string starts with an extended prefix we would need to remove it from the path before we call GetFullPathName as
        //    // it doesn't root extended paths correctly. We don't currently resolve extended paths, so we'll just assert here.
        //    Debug.Assert(IsPartiallyQualified(path) || !IsExtended(new string(path.ToArray())));

        //    uint result;
            
        //    //GetFullPathNameW(ref MemoryMarshal.GetReference(path), (uint)builder.Capacity, ref builder.GetPinnableReference(), IntPtr.Zero))

        //    while ((result = GetFullPathNameW(ref MemoryMarshal.GetReference(path), (uint)builder.Capacity, ref builder.GetPinnableReference(), IntPtr.Zero)) > builder.Capacity)
        //    {
        //        // Reported _size is greater than the buffer _size. Increase the capacity.
        //        builder.EnsureCapacity(checked((int)result));
        //    }

        //    if (result == 0)
        //    {
        //        // Failure, get the error and throw
        //        int errorCode = GetLastPInvokeError();
        //        if (errorCode == 0)
        //            errorCode = ERROR_BAD_PATHNAME;
        //        throw GetExceptionForWin32Error(errorCode, path.ToString());
        //    }

        //    builder.Length = (int)result;
        //}

        /// <summary>
        /// Normalize the given path.
        /// </summary>
        /// <remarks>
        /// Normalizes via Win32 GetFullPathName().
        /// </remarks>
        /// <param name="path">Path to normalize</param>
        /// <exception cref="PathTooLongException">Thrown if we have a string that is too large to fit into a UNICODE_STRING.</exception>
        /// <exception cref="IOException">Thrown if the path is empty.</exception>
        /// <returns>Normalized path</returns>
        internal static string Normalize(string path)
        {
            //var builder = new ValueStringBuilder(stackalloc char[MaxShortPath]);

            // Get the full path
            //GetFullPathName(path.AsSpan(), ref builder);
            //string result = null;
            //GetFullPathName(path, ref result);
            return Path.GetFullPath(path);

            // If we have the exact same string we were passed in, don't allocate another string.
            // TryExpandShortName does this input identity check.
            //string result = builder.AsSpan().IndexOf('~') >= 0
            //    ? TryExpandShortFileName(ref builder, originalPath: path)
            //    : builder.AsSpan().Equals(path.AsSpan(), StringComparison.Ordinal) ? path : builder.ToString();

            //// Clear the buffer
            //builder.Dispose();
            //return result;
        }

        internal static bool IsApp64Bit() 
        { 
            return (64 == (IntPtr.Size * 8));
        }

        internal static string TryExpandShortFileName(ref ValueStringBuilder outputBuilder, string? originalPath)
        {
            // We guarantee we'll expand short names for paths that only partially exist. As such, we need to find the part of the path that actually does exist. To
            // avoid allocating a lot we'll create only one input array and modify the contents with embedded nulls.

            Debug.Assert(!IsPartiallyQualified(outputBuilder.AsSpan()), "should have resolved by now");

            // We'll have one of a few cases by now (the normalized path will have already:
            //
            //  1. Dos path (C:\)
            //  2. Dos UNC (\\Server\Share)
            //  3. Dos device path (\\.\C:\, \\?\C:\)
            //
            // We want to put the extended syntax on the front if it doesn't already have it (for long path support and speed), which may mean switching from \\.\.
            //
            // Note that we will never get \??\ here as GetFullPathName() does not recognize \??\ and will return it as C:\??\ (or whatever the current drive is).

            int rootLength = GetRootLength(new string(outputBuilder.AsSpan().ToArray()));
            bool isDevice = IsDevice(new string(outputBuilder.AsSpan().ToArray()));

            // As this is a corner case we're not going to add a stackalloc here to keep the stack pressure down.
            ValueStringBuilder inputBuilder = default;

            bool isDosUnc = false;
            int rootDifference = 0;
            bool wasDotDevice = false;

            // Add the extended prefix before expanding to allow growth over MAX_PATH
            if (isDevice)
            {
                // We have one of the following (\\?\ or \\.\)
                inputBuilder.Append(outputBuilder.AsSpan());

                if (outputBuilder[2] == '.')
                {
                    wasDotDevice = true;
                    inputBuilder[2] = '?';
                }
            }
            else
            {
                isDosUnc = !IsDevice(new string(outputBuilder.AsSpan().ToArray())) && outputBuilder.Length > 1 && outputBuilder[0] == '\\' && outputBuilder[1] == '\\';
                rootDifference = PrependDevicePathChars(ref outputBuilder, isDosUnc, ref inputBuilder);
            }

            rootLength += rootDifference;
            int inputLength = inputBuilder.Length;

            bool success = false;
            int foundIndex = inputBuilder.Length - 1;

            while (!success)
            {
                uint result = GetLongPathNameW(
                    ref inputBuilder.GetPinnableReference(terminate: true), ref outputBuilder.GetPinnableReference(), (uint)outputBuilder.Capacity);

                // Replace any temporary null we added
                if (inputBuilder[foundIndex] == '\0') inputBuilder[foundIndex] = '\\';

                if (result == 0)
                {
                    // Look to see if we couldn't find the file
                    int error = GetLastPInvokeError();
                    if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
                    {
                        // Some other failure, give up
                        break;
                    }

                    // We couldn't find the path at the given index, start looking further back in the string.
                    foundIndex--;

                    for (; foundIndex > rootLength && inputBuilder[foundIndex] != '\\'; foundIndex--) ;
                    if (foundIndex == rootLength)
                    {
                        // Can't trim the path back any further
                        break;
                    }
                    else
                    {
                        // Temporarily set a null in the string to get Windows to look further up the path
                        inputBuilder[foundIndex] = '\0';
                    }
                }
                else if (result > outputBuilder.Capacity)
                {
                    // Not enough space. The result count for this API does not include the null terminator.
                    outputBuilder.EnsureCapacity(checked((int)result));
                }
                else
                {
                    // Found the path
                    success = true;
                    outputBuilder.Length = checked((int)result);
                    if (foundIndex < inputLength - 1)
                    {
                        // It was a partial find, put the non-existent part of the path back
                        outputBuilder.Append(inputBuilder.AsSpan(foundIndex, inputBuilder.Length - foundIndex));
                    }
                }
            }

            // If we were able to expand the path, use it, otherwise use the original full path result
            ref ValueStringBuilder builderToUse = ref (success ? ref outputBuilder : ref inputBuilder);

            // Switch back from \\?\ to \\.\ if necessary
            if (wasDotDevice)
                builderToUse[2] = '.';

            // Change from \\?\UNC\ to \\?\UN\\ if needed
            if (isDosUnc)
                builderToUse[UncExtendedPrefixLength - UncPrefixLength] = '\\';

            // Strip out any added characters at the front of the string
            ReadOnlySpan<char> output = builderToUse.AsSpan(rootDifference);

            string returnValue = ((originalPath != null) && output.Equals(originalPath.AsSpan(), StringComparison.Ordinal))
                ? originalPath : output.ToString();

            inputBuilder.Dispose();
            return returnValue;
        }

        // \\
        internal const int UncPrefixLength = 2;
        // \\?\UNC\, \\.\UNC\
        internal const int UncExtendedPrefixLength = 8;
        // \\?\, \\.\, \??\
        internal const int DevicePrefixLength = 4;
        internal const string UncExtendedPrefixToInsert = @"?\UNC\";
        internal const string UncExtendedPathPrefix = @"\\?\UNC\";
        internal const string UncNTPathPrefix = @"\??\UNC\";
        internal const string ExtendedPathPrefix = @"\\?\";

        internal static int PrependDevicePathChars(ref ValueStringBuilder content, bool isDosUnc, ref ValueStringBuilder buffer)
        {
            int length = content.Length;

            length += isDosUnc
                ? UncExtendedPrefixLength - UncPrefixLength
                : DevicePrefixLength;

            buffer.EnsureCapacity(length + 1);
            buffer.Length = 0;

            if (isDosUnc)
            {
                // Is a \\Server\Share, put \\?\UNC\ in the front
                buffer.Append(UncExtendedPathPrefix);

                // Copy Server\Share\... over to the buffer
                buffer.Append(content.AsSpan(UncPrefixLength));

                // Return the prefix difference
                return UncExtendedPrefixLength - UncPrefixLength;
            }
            else
            {
                // Not an UNC, put the \\?\ prefix in front, then the original string
                buffer.Append(ExtendedPathPrefix);
                buffer.Append(content.AsSpan());
                return DevicePrefixLength;
            }
        }

        [SecuritySafeCritical]
        internal unsafe static int GetRootLength(string path)
        {
            fixed (char* path2 = path)
            {
                return (int)GetRootLength(path2, (ulong)path.Length);
            }
        }
                
        [SecurityCritical]
        private unsafe static uint GetRootLength(char* path, ulong pathLength)
        {
            uint num = 0u;
            uint num2 = 2u;
            uint num3 = 2u;
            bool flag = StartsWithOrdinal(path, pathLength, "\\\\?\\");
            bool flag2 = StartsWithOrdinal(path, pathLength, "\\\\?\\UNC\\");
            if (flag)
            {
                if (flag2)
                {
                    num3 = (uint)"\\\\?\\UNC\\".Length;
                }
                else
                {
                    num2 += (uint)"\\\\?\\".Length;
                }
            }

            if ((!flag || flag2) && pathLength != 0 && IsDirectorySeparator(*path))
            {
                num = 1u;
                if (flag2 || (pathLength > 1 && IsDirectorySeparator(path[1])))
                {
                    num = num3;
                    int num4 = 2;
                    for (; num < pathLength; num++)
                    {
                        if (IsDirectorySeparator(path[num]) && --num4 <= 0)
                        {
                            break;
                        }
                    }
                }
            }
            else if (pathLength >= num2 && path[num2 - 1] == Path.VolumeSeparatorChar)
            {
                num = num2;
                if (pathLength >= num2 + 1 && IsDirectorySeparator(path[num2]))
                {
                    num++;
                }
            }

            return num;
        }

        [SecurityCritical]
        private unsafe static bool StartsWithOrdinal(char* source, ulong sourceLength, string value)
        {
            if (sourceLength < (ulong)value.Length)
            {
                return false;
            }

            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] != source[i])
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Try to remove relative segments from the given path (without combining with a root).
        /// </summary>
        /// <param name="path">Input path</param>
        /// <param name="rootLength">The length of the root of the given path</param>
        internal static string RemoveRelativeSegments(string path, int rootLength)
        {
            var sb = new ValueStringBuilder(stackalloc char[260 /* PathInternal.MaxShortPath */]);

            if (RemoveRelativeSegments(path.AsSpan(), rootLength, ref sb))
            {
                path = sb.ToString();
            }

            sb.Dispose();
            return path;
        }

        internal const char DirectorySeparatorChar = '\\';
        internal const char AltDirectorySeparatorChar = '/';

        /// <summary>
        /// Try to remove relative segments from the given path (without combining with a root).
        /// </summary>
        /// <param name="path">Input path</param>
        /// <param name="rootLength">The length of the root of the given path</param>
        /// <param name="sb">String builder that will store the result</param>
        /// <returns>"true" if the path was modified</returns>
        internal static bool RemoveRelativeSegments(ReadOnlySpan<char> path, int rootLength, ref ValueStringBuilder sb)
        {
            Debug.Assert(rootLength > 0);
            bool flippedSeparator = false;

            int skip = rootLength;
            // We treat "\.." , "\." and "\\" as a relative segment. We want to collapse the first separator past the root presuming
            // the root actually ends in a separator. Otherwise the first segment for RemoveRelativeSegments
            // in cases like "\\?\C:\.\" and "\\?\C:\..\", the first segment after the root will be ".\" and "..\" which is not considered as a relative segment and hence not be removed.
            if (IsDirectorySeparator(path[skip - 1]))
                skip--;

            // Remove "//", "/./", and "/../" from the path by copying each character to the output,
            // except the ones we're removing, such that the builder contains the normalized path
            // at the end.
            if (skip > 0)
            {
                sb.Append(path.Slice(0, skip));
            }

            for (int i = skip; i < path.Length; i++)
            {
                char c = path[i];

                if (IsDirectorySeparator(c) && i + 1 < path.Length)
                {
                    // Skip this character if it's a directory separator and if the next character is, too,
                    // e.g. "parent//child" => "parent/child"
                    if (IsDirectorySeparator(path[i + 1]))
                    {
                        continue;
                    }

                    // Skip this character and the next if it's referring to the current directory,
                    // e.g. "parent/./child" => "parent/child"
                    if ((i + 2 == path.Length || IsDirectorySeparator(path[i + 2])) &&
                        path[i + 1] == '.')
                    {
                        i++;
                        continue;
                    }

                    // Skip this character and the next two if it's referring to the parent directory,
                    // e.g. "parent/child/../grandchild" => "parent/grandchild"
                    if (i + 2 < path.Length &&
                        (i + 3 == path.Length || IsDirectorySeparator(path[i + 3])) &&
                        path[i + 1] == '.' && path[i + 2] == '.')
                    {
                        // Unwind back to the last slash (and if there isn't one, clear out everything).
                        int s;
                        for (s = sb.Length - 1; s >= skip; s--)
                        {
                            if (IsDirectorySeparator(sb[s]))
                            {
                                sb.Length = (i + 3 >= path.Length && s == skip) ? s + 1 : s; // to avoid removing the complete "\tmp\" segment in cases like \\?\C:\tmp\..\, C:\tmp\..
                                break;
                            }
                        }
                        if (s < skip)
                        {
                            sb.Length = skip;
                        }

                        i += 2;
                        continue;
                    }
                }

                // Normalize the directory separator if needed
                if (c != DirectorySeparatorChar && c == AltDirectorySeparatorChar)
                {
                    c = DirectorySeparatorChar;
                    flippedSeparator = true;
                }

                sb.Append(c);
            }

            // If we haven't changed the source path, return the original
            if (!flippedSeparator && sb.Length == path.Length)
            {
                return false;
            }

            // We may have eaten the trailing separator from the root when we started and not replaced it
            if (skip != rootLength && sb.Length < rootLength)
            {
                sb.Append(path[rootLength - 1]);
            }

            return true;
        }


        internal static ReadOnlySpan<char> GetVolumeName(ReadOnlySpan<char> path)
        {
            // 3 cases: UNC ("\\server\share"), Device ("\\?\C:\"), or Dos ("C:\")
            ReadOnlySpan<char> root = Path.GetPathRoot(new string(path.ToArray())).AsSpan();
            if (root.Length == 0)
                return root;

            // Cut from "\\?\UNC\Server\Share" to "Server\Share"
            // Cut from  "\\Server\Share" to "Server\Share"
            int startOffset = GetUncRootLength(path);
            if (startOffset == -1)
            {
                if (IsDevice(new string(path.ToArray())))
                {
                    startOffset = 4; // Cut from "\\?\C:\" to "C:"
                }
                else
                {
                    startOffset = 0; // e.g. "C:"
                }
            }

            ReadOnlySpan<char> pathToTrim = root.Slice(startOffset);
            return EndsInDirectorySeparator(pathToTrim) ? pathToTrim.Slice(0, pathToTrim.Length - 1) : pathToTrim;
        }

        internal static bool EndsInDirectorySeparator(ReadOnlySpan<char> path)
        {
            return (path[path.Length - 1] == Path.DirectorySeparatorChar);
        }

        internal static int GetUncRootLength(ReadOnlySpan<char> path)
        {
            const string UncExtendedPathPrefix = "\\\\?\\UNC\\";
            bool isDevice = IsDevice(new string(path.ToArray()));
            if (!isDevice && path.Slice(0, 2).Equals(@"\\".AsSpan(), StringComparison.Ordinal))
                return 2;
            else if (isDevice && path.Length >= 8
                && (path.Slice(0, 8).Equals(UncExtendedPathPrefix.AsSpan(),StringComparison.Ordinal)
                || path.Slice(5, 4).Equals(@"UNC\".AsSpan(), StringComparison.Ordinal)))
                return 8;

            return -1;
        }

        internal static bool IsDevice(string path)
        {
            if (!IsExtended(path))
            {
                if (path.Length >= 4 && IsDirectorySeparator(path[0]) && IsDirectorySeparator(path[1]) && (path[2] == '.' || path[2] == '?'))
                {
                    return IsDirectorySeparator(path[3]);
                }

                return false;
            }

            return true;
        }
                
        internal static bool IsExtended(string path)
        {
            if (path.Length >= 4 && path[0] == '\\' && (path[1] == '\\' || path[1] == '?') && path[2] == '?')
            {
                return path[3] == '\\';
            }

            return false;
        }

        internal static bool IsExtended(StringBuilder path)
        {
            if (path.Length >= 4 && path[0] == '\\' && (path[1] == '\\' || path[1] == '?') && path[2] == '?')
            {
                return path[3] == '\\';
            }

            return false;
        }
                

        public static ReadOnlySpan<char> Concat2(this ReadOnlySpan<char> first, ReadOnlySpan<char> second)
        {
            return new string(first.ToArray().Concat(second.ToArray()).ToArray()).AsSpan();
        }

        public static string Concat2(this string first, ReadOnlySpan<char> second)
        {
            return new string(first.ToArray().Concat(second.ToArray()).ToArray());
        }

        //private static unsafe string JoinInternal(ReadOnlySpan<char> first, ReadOnlySpan<char> second)
        //{
        //    Debug.Assert(first.Length > 0 && second.Length > 0, "should have dealt with empty paths");

        //    bool hasSeparator = IsDirectorySeparator(first[first.Length - 1])
        //        || IsDirectorySeparator(second[0]);

        //    fixed (char* f = &MemoryMarshal.GetReference(first), s = &MemoryMarshal.GetReference(second))
        //    {
        //        return string.Create(
        //            first.Length + second.Length + (hasSeparator ? 0 : 1),
        //            (First: (IntPtr)f, FirstLength: first.Length, Second: (IntPtr)s, SecondLength: second.Length, HasSeparator: hasSeparator),
        //            (destination, state) =>
        //            {
        //                new Span<char>((char*)state.First, state.FirstLength).CopyTo(destination);
        //                if (!state.HasSeparator)
        //                    destination[state.FirstLength] = PathInternal.DirectorySeparatorChar;
        //                new Span<char>((char*)state.Second, state.SecondLength).CopyTo(destination.Slice(state.FirstLength + (state.HasSeparator ? 0 : 1)));
        //            });
        //    }
        //}

        //public static string Create<TState>(int length, TState state, Action<char, TState> action)
        //{
        //    if (action == null)
        //        throw new ArgumentNullException(nameof(action));

        //    if (length <= 0)
        //    {
        //        if (length == 0)
        //            return String.Empty;
        //        throw new ArgumentOutOfRangeException(nameof(length));
        //    }

        //    string result = FastAllocateString(length);
        //    action(new Span<char>(ref result.GetRawStringData(), length), state);
        //    return result;
        //}

        ///// <summary>Creates a new string by using the specified provider to control the formatting of the specified interpolated string.</summary>
        ///// <param name="provider">An object that supplies culture-specific formatting information.</param>
        ///// <param name="handler">The interpolated string.</param>
        ///// <returns>The string that results for formatting the interpolated string using the specified format provider.</returns>
        //public static string Create(IFormatProvider? provider, [InterpolatedStringHandlerArgument(nameof(provider))] ref DefaultInterpolatedStringHandler handler) =>
        //    handler.ToStringAndClear();

        ///// <summary>Creates a new string by using the specified provider to control the formatting of the specified interpolated string.</summary>
        ///// <param name="provider">An object that supplies culture-specific formatting information.</param>
        ///// <param name="initialBuffer">The initial buffer that may be used as temporary space as part of the formatting operation. The contents of this buffer may be overwritten.</param>
        ///// <param name="handler">The interpolated string.</param>
        ///// <returns>The string that results for formatting the interpolated string using the specified format provider.</returns>
        //public static string Create(IFormatProvider? provider, Span<char> initialBuffer, [InterpolatedStringHandlerArgument("provider", "initialBuffer")] ref DefaultInterpolatedStringHandler handler) =>
        //    handler.ToStringAndClear();


        static public bool IsDirectorySeparator(char character) => character == Path.DirectorySeparatorChar;

        /// <summary>
        /// Decompiled from PathInternal
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal static bool IsValidDriveChar(char value)
        {
            if (value < 'A' || value > 'Z')
            {
                if (value >= 'a')
                {
                    return value <= 'z';
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Returns true if the path is fixed to a specific drive or UNC path. This method does no
        /// validation of the path (URIs will be returned as relative as a result).
        /// Returns false if the path specified is relative to the current drive or working directory.
        /// </summary>
        /// <remarks>
        /// Handles paths that use the alternate directory separator.  It is a frequent mistake to
        /// assume that rooted paths <see cref="Path.IsPathRooted(string)"/> are not relative.  This isn't the case.
        /// "C:a" is drive relative- meaning that it will be resolved against the current directory
        /// for C: (rooted, but relative). "C:\a" is rooted and not relative (the current directory
        /// will not be used to modify the path).
        /// </remarks>
        /// <exception cref="ArgumentNullException">
        /// Thrown if <paramref name="path"/> is null.
        /// </exception>
        public static bool IsPathFullyQualified(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            return IsPathFullyQualified(path.AsSpan());
        }

        public static bool IsPathFullyQualified(ReadOnlySpan<char> path)
        {
            return !IsPartiallyQualified(path);
        }

        static public string NormalizeFolderpath(string filepath)
        {
            //string result = System.IO.Path.GetFullPath(filepath).ToLowerInvariant();
            string result = filepath;

            if (Path.GetExtension(result)=="")
            {
                result = result.TrimEnd(new[] { '\\' });
                result += '\\';
            }
            else
            {
                result = Path.GetDirectoryName(result) + '\\';
            }
            return result;
        }

        #endregion
    }

    internal static class SR
    {

        /// <summary>Enum value was out of legal range.</summary>
        internal static string @ArgumentOutOfRange_Enum => GetResourceString("ArgumentOutOfRange_Enum", @"Enum value was out of legal range.");
        /// <summary>Non-negative number required.</summary>
        internal static string @ArgumentOutOfRange_NeedNonNegNum => GetResourceString("ArgumentOutOfRange_NeedNonNegNum", @"Non-negative number required.");
        /// <summary>Positive number required.</summary>
        internal static string @ArgumentOutOfRange_NeedPosNum => GetResourceString("ArgumentOutOfRange_NeedPosNum", @"Positive number required.");
        /// <summary>Empty name is not legal.</summary>
        internal static string @Argument_EmptyName => GetResourceString("Argument_EmptyName", @"Empty name is not legal.");
        /// <summary>The initial count for the semaphore must be greater than or equal to zero and less than the maximum count.</summary>
        internal static string @Argument_SemaphoreInitialMaximum => GetResourceString("Argument_SemaphoreInitialMaximum", @"The initial count for the semaphore must be greater than or equal to zero and less than the maximum count.");
        /// <summary>The length of the name exceeds the maximum limit.</summary>
        internal static string @Argument_WaitHandleNameTooLong => GetResourceString("Argument_WaitHandleNameTooLong", @"The length of the name exceeds the maximum limit.");
        /// <summary>Cannot create '{0}' because a file or directory with the same name already exists.</summary>
        internal static string @IO_AlreadyExists_Name => GetResourceString("IO_AlreadyExists_Name", @"Cannot create '{0}' because a file or directory with the same name already exists.");
        /// <summary>The file '{0}' already exists.</summary>
        internal static string @IO_FileExists_Name => GetResourceString("IO_FileExists_Name", @"The file '{0}' already exists.");
        /// <summary>Unable to find the specified file.</summary>
        internal static string @IO_FileNotFound => GetResourceString("IO_FileNotFound", @"Unable to find the specified file.");
        /// <summary>Could not find file '{0}'.</summary>
        internal static string @IO_FileNotFound_FileName => GetResourceString("IO_FileNotFound_FileName", @"Could not find file '{0}'.");
        /// <summary>Could not find a part of the path.</summary>
        internal static string @IO_PathNotFound_NoPathName => GetResourceString("IO_PathNotFound_NoPathName", @"Could not find a part of the path.");
        /// <summary>Could not find a part of the path '{0}'.</summary>
        internal static string @IO_PathNotFound_Path => GetResourceString("IO_PathNotFound_Path", @"Could not find a part of the path '{0}'.");
        /// <summary>The specified file name or path is too long, or a component of the specified path is too long.</summary>
        internal static string @IO_PathTooLong => GetResourceString("IO_PathTooLong", @"The specified file name or path is too long, or a component of the specified path is too long.");
        /// <summary>The path '{0}' is too long, or a component of the specified path is too long.</summary>
        internal static string @IO_PathTooLong_Path => GetResourceString("IO_PathTooLong_Path", @"The path '{0}' is too long, or a component of the specified path is too long.");
        /// <summary>The process cannot access the file '{0}' because it is being used by another process.</summary>
        internal static string @IO_SharingViolation_File => GetResourceString("IO_SharingViolation_File", @"The process cannot access the file '{0}' because it is being used by another process.");
        /// <summary>The process cannot access the file because it is being used by another process.</summary>
        internal static string @IO_SharingViolation_NoFileName => GetResourceString("IO_SharingViolation_NoFileName", @"The process cannot access the file because it is being used by another process.");
        /// <summary>Access Control List (ACL) APIs are part of resource management on Windows and are not supported on this platform.</summary>
        internal static string @PlatformNotSupported_AccessControl => GetResourceString("PlatformNotSupported_AccessControl", @"Access Control List (ACL) APIs are part of resource management on Windows and are not supported on this platform.");
        /// <summary>A WaitHandle with system-wide name '{0}' cannot be created. A WaitHandle of a different type might have the same name.</summary>
        internal static string @Threading_WaitHandleCannotBeOpenedException_InvalidHandle => GetResourceString("Threading_WaitHandleCannotBeOpenedException_InvalidHandle", @"A WaitHandle with system-wide name '{0}' cannot be created. A WaitHandle of a different type might have the same name.");
        /// <summary>Access to the path is denied.</summary>
        internal static string @UnauthorizedAccess_IODenied_NoPathName => GetResourceString("UnauthorizedAccess_IODenied_NoPathName", @"Access to the path is denied.");
        /// <summary>Access to the path '{0}' is denied.</summary>
        internal static string @UnauthorizedAccess_IODenied_Path => GetResourceString("UnauthorizedAccess_IODenied_Path", @"Access to the path '{0}' is denied.");
        /// <summary>A WaitHandle with system-wide name '{0}' cannot be created. A WaitHandle of a different type might have the same name.</summary>
        internal static string @WaitHandleCannotBeOpenedException_InvalidHandle => GetResourceString("WaitHandleCannotBeOpenedException_InvalidHandle", @"A WaitHandle with system-wide name '{0}' cannot be created. A WaitHandle of a different type might have the same name.");


        private static global::System.Resources.ResourceManager s_resourceManager = s_resourceManager = new global::System.Resources.ResourceManager(typeof(SR));
        internal static global::System.Resources.ResourceManager ResourceManager => s_resourceManager;

        /// <summary>Failed to create system events window thread.</summary>
        internal static string @ErrorCreateSystemEvents => GetResourceString("ErrorCreateSystemEvents", @"Failed to create system events window thread.");
        /// <summary>Cannot create timer.</summary>
        internal static string @ErrorCreateTimer => GetResourceString("ErrorCreateTimer", @"Cannot create timer.");
        /// <summary>Cannot end timer.</summary>
        internal static string @ErrorKillTimer => GetResourceString("ErrorKillTimer", @"Cannot end timer.");
        /// <summary>'{1}' is not a valid value for '{0}'. '{0}' must be greater than {2}.</summary>
        internal static string @InvalidLowBoundArgument => GetResourceString("InvalidLowBoundArgument", @"'{1}' is not a valid value for '{0}'. '{0}' must be greater than {2}.");
        /// <summary>SystemEvents is not supported on this platform.</summary>
        internal static string @PlatformNotSupported_SystemEvents => GetResourceString("PlatformNotSupported_SystemEvents", @"SystemEvents is not supported on this platform.");

        private static readonly bool s_usingResourceKeys = AppContext.TryGetSwitch("System.Resources.UseSystemResourceKeys", out bool usingResourceKeys) ? usingResourceKeys : false;

        // This method is used to decide if we need to append the exception message parameters to the message when calling SR.Format.
        // by default it returns the value of System.Resources.UseSystemResourceKeys AppContext switch or false if not specified.
        // Native code generators can replace the value this returns based on user input at the time of native code generation.
        // The trimming tools are also capable of replacing the value of this method when the application is being trimmed.
        internal static bool UsingResourceKeys() => s_usingResourceKeys;

        internal static string GetResourceString(string resourceKey)
        {
            if (UsingResourceKeys())
            {
                return resourceKey;
            }

            string? resourceString = null;
            try
            {
                resourceString =
#if SYSTEM_PRIVATE_CORELIB || NATIVEAOT
                    InternalGetResourceString(resourceKey);
#else
                    ResourceManager.GetString(resourceKey);
#endif
            }
            catch (MissingManifestResourceException) { }

            return resourceString!; // only null if missing resources
        }

        internal static string GetResourceString(string resourceKey, string defaultString)
        {
            string resourceString = GetResourceString(resourceKey);

            return resourceKey == resourceString || resourceString == null ? defaultString : resourceString;
        }

        internal static string Format(string resourceFormat, object? p1)
        {
            if (UsingResourceKeys())
            {
                return string.Join(", ", resourceFormat, p1);
            }

            return string.Format(resourceFormat, p1);
        }

        internal static string Format(string resourceFormat, object? p1, object? p2)
        {
            if (UsingResourceKeys())
            {
                return string.Join(", ", resourceFormat, p1, p2);
            }

            return string.Format(resourceFormat, p1, p2);
        }

        internal static string Format(string resourceFormat, object? p1, object? p2, object? p3)
        {
            if (UsingResourceKeys())
            {
                return string.Join(", ", resourceFormat, p1, p2, p3);
            }

            return string.Format(resourceFormat, p1, p2, p3);
        }

        internal static string Format(string resourceFormat, params object?[]? args)
        {
            if (args != null)
            {
                if (UsingResourceKeys())
                {
                    return resourceFormat + ", " + string.Join(", ", args);
                }

                return string.Format(resourceFormat, args);
            }

            return resourceFormat;
        }

        internal static string Format(IFormatProvider? provider, string resourceFormat, object? p1)
        {
            if (UsingResourceKeys())
            {
                return string.Join(", ", resourceFormat, p1);
            }

            return string.Format(provider, resourceFormat, p1);
        }

        internal static string Format(IFormatProvider? provider, string resourceFormat, object? p1, object? p2)
        {
            if (UsingResourceKeys())
            {
                return string.Join(", ", resourceFormat, p1, p2);
            }

            return string.Format(provider, resourceFormat, p1, p2);
        }

        internal static string Format(IFormatProvider? provider, string resourceFormat, object? p1, object? p2, object? p3)
        {
            if (UsingResourceKeys())
            {
                return string.Join(", ", resourceFormat, p1, p2, p3);
            }

            return string.Format(provider, resourceFormat, p1, p2, p3);
        }

        internal static string Format(IFormatProvider? provider, string resourceFormat, params object?[]? args)
        {
            if (args != null)
            {
                if (UsingResourceKeys())
                {
                    return resourceFormat + ", " + string.Join(", ", args);
                }

                return string.Format(provider, resourceFormat, args);
            }

            return resourceFormat;
        }
    }
}
