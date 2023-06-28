/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Jun 22 15:08:23 2023
 */
/* Compiler settings for IHelloWorld.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_IHelloWorld = {0xA851A7FE,0x4903,0x48AF,{0xA6,0x94,0x51,0xFE,0xB7,0x55,0xEE,0x5B}};


const IID LIBID_HelloWorldLib = {0x9EBDD250,0x565C,0x4182,{0xB5,0xE9,0x70,0xCF,0x63,0xA8,0x96,0xE1}};


const CLSID CLSID_HelloWorld = {0xDC0F3891,0x93F3,0x42E9,{0xA1,0x17,0x72,0x9B,0x4F,0x3C,0x77,0x5A}};


#ifdef __cplusplus
}
#endif

