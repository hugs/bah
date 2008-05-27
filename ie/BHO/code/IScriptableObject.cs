using System;
using System.Runtime.InteropServices;

namespace BHOSeleniumIce 
{
    [
    InterfaceType(ComInterfaceType.InterfaceIsIDispatch)
    ]
    public interface IScriptableObject
    {
        void helloWorld();
        void showMessage(String textToShow);
        void setFileField(String fieldId, String filePath);
    }
}
