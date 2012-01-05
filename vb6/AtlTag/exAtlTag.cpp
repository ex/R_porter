// exAtlTag.cpp : Implementation of CexAtlTag

#include "stdafx.h"
#include "AtlTag.h"
#include "exAtlTag.h"

#include "string.h"

/////////////////////////////////////////////////////////////////////////////
// CexAtlTag
STDMETHODIMP CexAtlTag::SetPathFile(BSTR pathFile, LONG BufferLen, BYTE bVerifyBR, LONG cod)
{
    USES_CONVERSION;

	//------------------------------------------------------
	// solo cadenas de MAX_PATH caracteres
	memset (ms_Path, 0, sizeof(ms_Path));
	strncpy (ms_Path, (W2A(pathFile)), sizeof(ms_Path) - 1);

	mn_ErrorNumber = 0;

	if (cod == 97114029)
	{
		if (clxTag.SetBufferLegth (BufferLen))
		{
			if (clxTag.OpenFile (ms_Path))
			{
				if (clxTag.ReadFile())
				{
					if (clxTag.GetFrameHeader (bVerifyBR != 0))
					{
						if (clxTag.header.variable != 0)
						{
							mn_Bitrate = 0;
						}
						else
						{
							mn_Bitrate = clxTag.header.bitrate;
						}

						mn_Mode = clxTag.header.mode;
						mn_Layer = clxTag.header.layer;
						mn_Mpeg = clxTag.header.mpeg;
						mn_SampleRate = clxTag.header.samplerate;

						clxTag.CloseFile();
						clxTag.FreeBuffer(); 
						return S_OK;
					}

					clxTag.CloseFile();
					clxTag.FreeBuffer(); 
				}
			}
		}
	}

	mn_ErrorNumber = 1;

	return S_OK;
}


STDMETHODIMP CexAtlTag::get_Bitrate(INT *pVal)
{
	*pVal = mn_Bitrate;

	return S_OK;
}

STDMETHODIMP CexAtlTag::get_Mpeg(short *pVal)
{
	*pVal = mn_Mpeg;

	return S_OK;
}

STDMETHODIMP CexAtlTag::get_Layer(short *pVal)
{
	*pVal = mn_Layer;

	return S_OK;
}

STDMETHODIMP CexAtlTag::get_SampleRate(INT *pVal)
{
	*pVal = mn_SampleRate;

	return S_OK;
}

STDMETHODIMP CexAtlTag::get_Mode(short *pVal)
{
	*pVal = mn_Mode;

	return S_OK;
}

STDMETHODIMP CexAtlTag::get_ErrorNumber(short *pVal)
{
	*pVal = mn_ErrorNumber;

	return S_OK;
}

//---------------------------------------------------------------------------
// con este metodo se abre el buffer para SetPathFile2
//
STDMETHODIMP CexAtlTag::SetBufferLength(INT BufferLen)
{
	if (!clxTag.SetBufferLegth (BufferLen))
	{
		mn_ErrorNumber = 2;
	}

	return S_OK;
}

//---------------------------------------------------------------------------
// En este metodo no se cierra el buffer y se supone que ya esta abierto
//
STDMETHODIMP CexAtlTag::SetPathFile2(BSTR pathFile, BYTE bVerifyBR, LONG cod)
{
	USES_CONVERSION;
	
	//------------------------------------------------------
	// solo cadenas de MAX_PATH caracteres
	memset (ms_Path, 0, sizeof(ms_Path));
	strncpy (ms_Path, (W2A(pathFile)), sizeof(ms_Path) - 1);

	mn_ErrorNumber = 0;

	if (cod == 97114029)
	{
		if (clxTag.OpenFile (ms_Path))
		{
			if (clxTag.ReadFile())
			{
				if (clxTag.GetFrameHeader (bVerifyBR != 0))
				{
					if (clxTag.header.variable != 0)
					{
						mn_Bitrate = 0;
					}
					else
					{
						mn_Bitrate = clxTag.header.bitrate;
					}

					mn_Mode = clxTag.header.mode;
					mn_Layer = clxTag.header.layer;
					mn_Mpeg = clxTag.header.mpeg;
					mn_SampleRate = clxTag.header.samplerate;

					clxTag.CloseFile();
					return S_OK;
				}

				clxTag.CloseFile();
			}
		}
	}

	mn_ErrorNumber = 1;

	return S_OK;
}
//---------------------------------------------------------------------------
// con este metodo se cierra el buffer (aunque tambien lo hace el DTOR)
//
STDMETHODIMP CexAtlTag::FreeBuffer()
{
	clxTag.FreeBuffer(); 

	return S_OK;
}
