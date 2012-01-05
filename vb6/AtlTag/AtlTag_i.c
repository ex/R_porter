/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sun Mar 21 10:54:53 2004
 */
/* Compiler settings for C:\Esau\Dev\R_porter\AtlTag\AtlTag.idl:
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

const IID IID_IexAtlTag = {0xED693590,0x6D10,0x4270,{0xA7,0x78,0x45,0xD0,0x09,0x4C,0xBF,0x6D}};


const IID IID_IexAtlTag2 = {0xCDB2899F,0xAFDE,0x47f4,{0xB1,0x52,0x56,0xB3,0xCD,0x4C,0x63,0x3F}};


const IID LIBID_ATLTAGLib = {0xB85EE4CE,0x0C3F,0x423B,{0xA0,0xE8,0x96,0xC7,0x55,0xEE,0xFE,0x24}};


const CLSID CLSID_exAtlTag = {0xD020FF56,0x24F5,0x4F5E,{0xA7,0x76,0x7E,0xCC,0x9C,0xB8,0x72,0xED}};


#ifdef __cplusplus
}
#endif

