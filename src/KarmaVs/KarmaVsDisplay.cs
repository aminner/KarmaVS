using System;
using System.ComponentModel;
using System.Drawing;
using devcoach.Tools.Properties;

using System.Drawing.Imaging;

namespace devcoach.Tools
{
    sealed internal class KarmaVsDisplay
    {
        public static int KarmaErrors =0;
        private BackgroundWorker worker;
        private bool _enabled;
        public bool Enabled
        {
            get { return _enabled; }
            set
            {
                _enabled = value;
                KarmaVsPackage.SetMenuStatus();
                UpdateDisplay();
            }
        }

        private DisplayType _displayType; 
        public DisplayType Type
        {
            get { return _displayType; }
            set
            {
                _displayType = value;
                UpdateDisplay();
            }
        }
        private void UpdateDisplay()
        {
            if (worker == null)
            {
                worker = new BackgroundWorker();
                worker.DoWork += delegate { SetDisplay(); };
                worker.RunWorkerAsync();
            }
        }
        private Bitmap GetDisplayObject(DisplayType type)
        {
            switch (type)
            {
                case DisplayType.BuildError:
                    return Resources._1420674003_001_11;
                case DisplayType.Failure:
                    return Resources._1420678809_001_19;
                case DisplayType.Running:
                    return Resources._1420678868_001_25;
                case DisplayType.Success:
                    return Resources._1420678794_001_18;
                default:
                    return null;
            }
        }
        private void SetDisplay()
        {
            var currentStatus = _displayType;
            var tmp = new Bitmap(25, 25);
            object icon = tmp.GetHbitmap();
            while (Enabled)
            {
                try
                {
                    if (KarmaExecution._karmaProcess != null && DateTimeOffset.Now - KarmaExecution._processStart > TimeSpan.FromSeconds(1))
                    {
                        _displayType = (KarmaErrors> 0) ? DisplayType.Failure : DisplayType.Success;
                    }
                }
                catch (Exception) { }

                if (currentStatus != _displayType)
                {
                    if (_displayType == DisplayType.Running)
                    {
                        KarmaErrors = 0;
                    }
                    currentStatus = _displayType;
                    int frozen;
                    KarmaVsPackage.StatusBar.IsFrozen(out frozen);
                    if (frozen == 0)
                    {
                        KarmaVsPackage.StatusBar.Animation(0, ref icon);
                        var image = GetDisplayObject(currentStatus);
                        //var tmp = ConvertToRGB(image);
                        icon = ConvertToRGB(image).GetHbitmap();
                        KarmaVsPackage.StatusBar.Animation(1, ref icon);
                    }
                }
            }
            KarmaVsPackage.StatusBar.Animation(0, ref icon);
        }

        public enum DisplayType
        {
            Success,
            Failure,
            Running,
            BuildError
        };
        public static Bitmap ConvertToRGB(Bitmap original)
        {
            var shell5 = GetIVsUIShell5();

            var backgroundColor = shell5.GetThemedColor(new System.Guid("DC0AF70E-5097-4DD3-9983-5A98C3A19942"), "ToolWindowBackground", 0);
            //var outputBitmap = GetInvertedBitmap(shell5, inputBitmap, almostGreenColor, backgroundColor);
            Bitmap newImage = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb);
            newImage.SetResolution(original.HorizontalResolution,
                                   original.VerticalResolution);
            Graphics g = Graphics.FromImage(newImage);
            g.Clear(ColorTranslator.FromWin32(Convert.ToInt32(backgroundColor & 0xffffff)));
            g.DrawImageUnscaled(original, 0, 0);
            g.Dispose();
            return newImage;
        }
        private static Microsoft.VisualStudio.Shell.Interop.IVsUIShell5 GetIVsUIShell5()
        {

            Microsoft.VisualStudio.OLE.Interop.IServiceProvider sp = null;
            Type SVsUIShellType = null;
            int hr = 0;
            IntPtr serviceIntPtr;
            Microsoft.VisualStudio.Shell.Interop.IVsUIShell5 shell5 = null;
            object serviceObject = null;

            sp = (Microsoft.VisualStudio.OLE.Interop.IServiceProvider)KarmaVsPackage.Application;

            SVsUIShellType = typeof(Microsoft.VisualStudio.Shell.Interop.SVsUIShell);

            hr = sp.QueryService(SVsUIShellType.GUID, SVsUIShellType.GUID, out serviceIntPtr);

            if (hr != 0)
            {
                System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr);
            }
            else if (!serviceIntPtr.Equals(IntPtr.Zero))
            {
                serviceObject = System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(serviceIntPtr);

                shell5 = (Microsoft.VisualStudio.Shell.Interop.IVsUIShell5)serviceObject;

                System.Runtime.InteropServices.Marshal.Release(serviceIntPtr);
            }
            return shell5;
        }
    }
}
