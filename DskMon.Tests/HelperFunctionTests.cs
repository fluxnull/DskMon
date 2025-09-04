using System;
using System.Reflection;
using DskMon;
using Xunit;

public class HelperFunctionTests
{
    private static MethodInfo GetPrivateMethod(string methodName)
    {
        var method = typeof(DskMon.DskMon).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static);
        if (method == null)
        {
            throw new Exception($"Could not find private method '{methodName}' on type 'DskMon.DskMon'");
        }
        return method;
    }

    [Theory]
    [InlineData("USBSTOR\\DISK&VEN_G-DRIVE&PROD_MOBILE&REV_1019\\575831314133343935353134&0", "575831314133343935353134")]
    [InlineData("SCSI\\DISK&VEN_WDC&PROD_WD10EZEX-00BN5A0\\4&1B8D4B6&0&000000", "4&1B8D4B6&0&000000")]
    [InlineData("IDE\\DISKWDC_WD10EZEX-00BN5A0__________________WD-WCC3F0XJ5Y5E\\4&1B8D4B6&0&0.0.0", "4&1B8D4B6&0&0.0.0")]
    [InlineData("", "")]
    [InlineData(null, "")]
    public void ParseSerialFromPnP_ReturnsCorrectSerial(string pnpDeviceId, string expectedSerial)
    {
        var method = GetPrivateMethod("ParseSerialFromPnP");
        var result = (string)method.Invoke(null, new object[] { pnpDeviceId });
        Assert.Equal(expectedSerial, result);
    }

    [Theory]
    [InlineData("USBSTOR\\DISK&VEN_G-DRIVE&PROD_MOBILE&REV_1019\\575831314133343935353134&0", "G-DRIVE")]
    [InlineData("SCSI\\DISK&VEN_WDC&PROD_WD10EZEX-00BN5A0\\4&1B8D4B6&0&000000", "WDC")]
    [InlineData("IDE\\DISKWDC_WD10EZEX-00BN5A0__________________WD-WCC3F0XJ5Y5E\\4&1B8D4B6&0&0.0.0", "")]
    [InlineData("", "")]
    [InlineData(null, "")]
    public void VendorFromPnP_ReturnsCorrectVendor(string pnpDeviceId, string expectedVendor)
    {
        var method = GetPrivateMethod("VendorFromPnP");
        var result = (string)method.Invoke(null, new object[] { pnpDeviceId });
        Assert.Equal(expectedVendor, result);
    }
}
