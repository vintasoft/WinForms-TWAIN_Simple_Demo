using System;
using System.ComponentModel;
using System.Windows.Forms;
using Vintasoft.Twain;

namespace TwainSimpleDemo
{
    public partial class MainForm : Form
    {

        #region Fields

        Device _device;

        #endregion



        #region Constructor

        public MainForm()
        {
            InitializeComponent();

            this.Text = string.Format("VintaSoft TWAIN Simple Demo v{0}", TwainGlobalSettings.ProductVersion);
        }

        #endregion



        #region Methods

        /// <summary>
        /// Scans images.
        /// </summary>
        private void scanImagesButton_Click(object sender, EventArgs e)
        {
            try
            {
                // disable application UI
                scanImagesButton.Enabled = false;

                // create TWAIN device manager
                using (DeviceManager deviceManager = new DeviceManager(this, this.Handle))
                {
                    try
                    {
                        // try to find TWAIN device manager
                        deviceManager.IsTwain2Compatible = twain2CheckBox.Checked;
                    }
                    catch (Exception ex)
                    {
                        // show dialog with error message
                        MessageBox.Show(GetFullExceptionMessage(ex), "TWAIN device manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // if 64-bit TWAIN2 device manager is used
                    if (IntPtr.Size == 8 && deviceManager.IsTwain2Compatible)
                    {
                        if (!InitTwain2DeviceManagerMode(deviceManager))
                            return;
                    }

                    try
                    {
                        // open the device manager
                        deviceManager.Open();
                    }
                    catch (Exception ex)
                    {
                        // show dialog with error message
                        MessageBox.Show(GetFullExceptionMessage(ex), "TWAIN device manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // if devices are NOT found
                    if (deviceManager.Devices.Count == 0)
                    {
                        MessageBox.Show("Devices are not found.");
                        return;
                    }

                    // if device is NOT selected
                    if (!deviceManager.ShowDefaultDeviceSelectionDialog())
                    {
                        MessageBox.Show("Device is not selected.");
                        return;
                    }

                    // get reference to the selected device
                    _device = deviceManager.DefaultDevice;

                    // set scan settings
                    _device.ShowUI = showUiCheckBox.Checked;
                    _device.ShowIndicators = showIndicatorsCheckBox.Checked;
                    _device.DisableAfterAcquire = !_device.ShowUI;
                    _device.CloseAfterModalAcquire = false;

                    AcquireModalState acquireModalState;
                    do
                    {
                        // synchronously acquire image from device
                        acquireModalState = _device.AcquireModal();
                        switch (acquireModalState)
                        {
                            case AcquireModalState.ImageAcquired:
                                // dispose previous bitmap in the picture box
                                if (pictureBox1.Image != null)
                                {
                                    pictureBox1.Image.Dispose();
                                    pictureBox1.Image = null;
                                }

                                // set a bitmap in the picture box
                                pictureBox1.Image = _device.AcquiredImage.GetAsBitmap(true);

                                // dispose an acquired image
                                _device.AcquiredImage.Dispose();
                                break;

                            case AcquireModalState.ScanCanceled:
                                MessageBox.Show("Scan is canceled.");
                                break;

                            case AcquireModalState.ScanFailed:
                                MessageBox.Show(string.Format("Scan is failed: {0}", _device.ErrorString));
                                break;
                        }
                    }
                    while (acquireModalState != AcquireModalState.None);

                    // close the device
                    _device.Close();
                    _device = null;

                    // close the device manager
                    deviceManager.Close();
                }
            }
            catch (TwainException ex)
            {
                MessageBox.Show(GetFullExceptionMessage(ex));
            }
            catch (Exception ex)
            {
                System.ComponentModel.LicenseException licenseException = GetLicenseException(ex);
                if (licenseException != null)
                {
                    // show information about licensing exception
                    MessageBox.Show(string.Format("{0}: {1}", licenseException.GetType().Name, licenseException.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    string[] dirs = new string[] { ".", "..", @"..\..\", @"..\..\..\", @"..\..\..\..\..\", @"..\..\..\..\..\..\..\" };
                    // for each directory
                    for (int i = 0; i < dirs.Length; i++)
                    {
                        string filename = System.IO.Path.Combine(dirs[i], "VSTwainNetEvaluationLicenseManager.exe");
                        // if VintaSoft Evaluation License Manager exists in directory
                        if (System.IO.File.Exists(filename))
                        {
                            // start Vintasoft Evaluation License Manager for getting the evaluation license
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            process.StartInfo.FileName = filename;
                            process.Start();
                        }
                    }
                }
                else
                {
                    throw;
                }
            }
            finally
            {
                // enable application UI
                scanImagesButton.Enabled = true;
            }
        }

        /// <summary>
        /// Initializes the device manager mode.
        /// </summary>
        /// <param name="deviceManager">The TWAIN device manager.</param>
        private bool InitTwain2DeviceManagerMode(DeviceManager deviceManager)
        {
            // create a form that allows to view and edit mode of 64-bit TWAIN2 device manager
            using (SelectDeviceManagerModeForm form = new SelectDeviceManagerModeForm())
            {
                // initialize form
                form.StartPosition = FormStartPosition.CenterParent;
                form.Owner = this;
                form.Use32BitDevices = deviceManager.Are32BitDevicesUsed;

                // show dialog
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // if device manager mode is changed
                    if (form.Use32BitDevices != deviceManager.Are32BitDevicesUsed)
                    {
                        try
                        {
                            // if 32-bit devices must be used
                            if (form.Use32BitDevices)
                                deviceManager.Use32BitDevices();
                            else
                                deviceManager.Use64BitDevices();
                        }
                        catch (TwainDeviceManagerException ex)
                        {
                            // show dialog with error message
                            MessageBox.Show(GetFullExceptionMessage(ex), "TWAIN device manager", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            return false;
                        }
                    }
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_device != null)
            {
                if (_device.State != DeviceState.Closed)
                    _device.Close();
            }
        }

        /// <summary>
        /// Returns the message of exception and inner exceptions.
        /// </summary>
        private string GetFullExceptionMessage(Exception ex)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(ex.Message);

            Exception innerException = ex.InnerException;
            while (innerException != null)
            {
                if (ex.Message != innerException.Message)
                    sb.AppendLine(string.Format("Inner exception: {0}", innerException.Message));
                innerException = innerException.InnerException;
            }

            return sb.ToString();
        }

        /// <summary>
        /// Returns the license exception from specified exception.
        /// </summary>
        /// <param name="exceptionObject">The exception object.</param>
        /// <returns>Instance of <see cref="LicenseException"/>.</returns>
        private static System.ComponentModel.LicenseException GetLicenseException(object exceptionObject)
        {
            Exception ex = exceptionObject as Exception;
            if (ex == null)
                return null;
            if (ex is System.ComponentModel.LicenseException)
                return (System.ComponentModel.LicenseException)exceptionObject;
            if (ex.InnerException != null)
                return GetLicenseException(ex.InnerException);
            return null;
        }

        #endregion

    }
}