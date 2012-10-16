load_cpp_dll_in_excel_with_vb
=============================

As the name say. Just a small tutorial how to create a c++ dll compatible with excel

You need 2 projects : 
* 1 Visual Studio Win32 DLL project
* 1 Excel Worksheet

Visual Studio
-------------
The settings for the project are all default (generated by Visual Studio) except :
* Linker
 * Input
  * Module Definition File : extern.def
 * Advanced
  * Entry Point : DllMain

The C++ source code contains at least 3 functions.
* The function you want to create
* The entry point of the DLL DllMain needed by windows DLL system.
* A Dummy function which is a void function taking no parameters and doing nothing. It is our trick to make Excel load the dll if it is in the same directory. Otherwise, except failure of loading it...

An example code is in **dll.cpp**

Excel
-----
For Excel, we need a little trick for loading the dll from the same folder of the .xls file.
The trick consist of executing a dummy function defined in the dll at the .xls loading

The dummy system is written in **dummy.vb**
The normal call of the function in vb is in **function_call.vb**

C++ and Excel Types compatibility
---------------------------------
<table>
  <tr>
    <th>Excel/VBA</th><th>C++</th>
  </tr>
  <tr>
    <td>Byte</td><td>unsigned char</td>
  </tr>
  <tr>
    <td>Boolean</td><td>[signed] short</td>
  </tr>
  <tr>
    <td>Long</td><td>[signed] [long] int</td>
  </tr>
  <tr>
    <td>Currency</td><td>CY (type de &lt;wtypes.h&gt;)</td>
  </tr>
  <tr>
    <td>Single</td><td>float</td>
  </tr>
  <tr>
    <td>Double</td><td>double</td>
  </tr>
  <tr>
    <td>Date</td><td>double<br/>DATE (type de &lt;wtypes.h&gt;)</td>
  </tr>
  <tr>
    <td>String</td><td>BSTR (type de &lt;wtypes.h&gt;)</td>
  </tr>
  <tr>
    <td>Variant</td><td>VARIANT (type de &lt;oaidl.h&gt;)</td>
  </tr>
</table>

Initialize a BSTR in C++
________________________
    char *mess="Hello World";
    BSTR ret=SysAllocStringByteLen(mess,strlen(mess));

Export an unsigned long long (ULONG64) from C++ to VB and back
______________________________________________________________
    unsigned long long X = 0x0123456789abcdef
    unsigned int x1 = (X & (0xffffffffLL << 32)) >> 32;
    unsigned int x2 = X & 0xffffffffLL;

    unsigned long long reconstructed;
    reconstructed = x1;
    reconstructed <<= 32;
    reconstructed |= x2;
