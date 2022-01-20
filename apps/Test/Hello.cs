using System;
using System.Runtime.InteropServices;
using Kompas6API5;
using ComHelpers;

namespace Test
{
    // ReSharper disable once UnusedType.Global
    // ReSharper disable once ClassNeverInstantiated.Global
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Hello
    {
        // ReSharper disable once UnusedMember.Global
        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "Hello - the first kompas app";
        }

        /// <summary>
        /// Main app entry
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public void ExternalRunCommand(
            // ReSharper disable once UnusedParameter.Global
            [In] short command,
            // ReSharper disable once UnusedParameter.Global
            [In] short mode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object kompasObj)
        {
            KompasObject kompas = (KompasObject) kompasObj;
            kompas.ksMessage("Hello Kompas!");
        }

        #region COM Registration
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            new Registrator(t).RegisterKompasLib();
        }
		
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            new Registrator(t).UnregisterKompasLib();
        }
        #endregion
    }
}