//*******************************************************************
// ExTag.cpp: implementation of the CExTag class.
//-------------------------------------------------------------------
// Proposito	:	Extract information from MPEG file (MP3)
// Creado por	:	Esau R.O.
// Fecha		:	Agosto 2003
//*******************************************************************

#include "ExTag.h"


#include <fcntl.h>
#include <io.h>
#include <stdio.h>
#include <stdlib.h>


unsigned int const bitrate_table[5][15] = {
  /* MPEG-1 */
  { 0,  32,  40,  48,  56,  64,  80,  96, 112, 128, 160, 192, 224, 256, 320 },	/* III		*/
  { 0,  32,  48,  56,  64,  80,  96, 112, 128, 160, 192, 224, 256, 320, 384 },	/* II		*/
  { 0,  32,  64,  96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448 },	/* I		*/

  /* MPEG-2 LSF */
  { 0,   8,  16,  24,  32,  40,  48,  56,  64,  80,  96, 112, 128, 144, 160 },	/* II & III */
  { 0,  32,  48,  56,  64,  80,  96, 112, 128, 144, 160, 176, 192, 224, 256 }	/* I		*/
};


unsigned int const samplerate_table[3] = { 44100, 48000, 32000 };



//*******************************************************************
// Construction/Destruction
//*******************************************************************
CExTag::CExTag()
{
	pBuffer = NULL;
	ml_BufferLen = 0;
	mb_FileOpen = false;
}


CExTag::~CExTag()
{
	if (NULL != pBuffer)
	{
		free (pBuffer);
		pBuffer = NULL;
	}

	if (mb_FileOpen)
	{
		close (nFile);
		mb_FileOpen = false;
	}
}


//*******************************************************************
// member functions
//*******************************************************************
bool CExTag::OpenFile (char *szPath)
{
	if ((strlen(szPath) <= MAX_PATH) && !mb_FileOpen)
	{
		strcpy (ms_PathFile, szPath);

		nFile = open (szPath, O_RDONLY | O_BINARY);

		if (nFile == -1)
		{
			return false;
		}
		else
		{
			mb_FileOpen = true;
			return true;
		}
	}
	else
		return false;
}

bool CExTag::SetBufferLegth (long length)
{
	if ((length > 0) && (NULL == pBuffer) && (0 == ml_BufferLen))
	{
		pBuffer = (unsigned char *) malloc (length);
		
		if (NULL != pBuffer)
		{
			ml_BufferLen = length;
			return true;
		}

		return false;
	}
	else
		return false;
}

bool CExTag::ReadFile()
{
	if ((NULL != pBuffer) && mb_FileOpen)
	{
		if (read (nFile, pBuffer, ml_BufferLen) <= 0)
		{
			return false;
		}

		return true;
	}

	return false;
}

bool CExTag::GetFrameHeader(bool bVerifyBitrate)
{
	if (NULL != pBuffer)
	{
		unsigned char	byte1, byte2;
		unsigned int	index;
		long			k = 0;
		
		while (k < ml_BufferLen - 1)
		{
			byte1 = *(pBuffer + k);
			byte2 = *(pBuffer + k + 1);
			
			if ((byte1 == 0xFF) && ((byte2 & 0xE0) == 0xE0))
			{
				// parece que encontramos el header...

				//------------------------------------
				// MPEG
				header.mpeg = ((byte2 & 0x18) >> 3);

				if (header.mpeg == EX_MPEG_FAIL)
				{
					goto BAD_HEADER;
				}

				//------------------------------------
				// LAYER
				header.layer = ((byte2 & 0x06) >> 1);

				if (header.layer == EX_LAYER_FAIL)
				{
					goto BAD_HEADER;
				}
				
				//------------------------------------
				// protection bit 
				header.crc = (byte2 & 0x01);

				byte2 = *(pBuffer + k + 2);

				//------------------------------------
				// bitrate_index
				index = ((byte2 & 0xF0) >> 4);

				if (index == 15)
				{
					goto BAD_HEADER;
				}
				else
				{
					if (header.mpeg == EX_MPEG_1)
					{
						header.bitrate = bitrate_table[header.layer - 1][index];
					}
					else
					{
						header.bitrate = bitrate_table[3 + ((header.layer - 1) >> 1)][index];
					}
				}

				//------------------------------------
				// sampling_frequency
				index = ((byte2 & 0x0C) >> 2);

				if (index == 3)
				{
					goto BAD_HEADER;
				}
				else
				{
					header.samplerate = samplerate_table[index];

					if (header.mpeg == EX_MPEG_2)
					{
						header.samplerate /= 2;
					}

					if (header.mpeg == EX_MPEG_2_5)
					{
						header.samplerate /= 4;
					}
				}

				//------------------------------------
				// padding_bit
				header.padding = ((byte2 & 0x02) >> 1);

				//------------------------------------
				// private_bit (ignored)

				byte2 = *(pBuffer + k + 3);

				//------------------------------------
				// channel mode
				header.mode = ((byte2 & 0xC0) >> 6);

				//------------------------------------
				// mode_extension (ignored)

				//------------------------------------
				// copyright
				header.copyright = ((byte2 & 0x08) >> 3);

				//------------------------------------
				// original/copy
				header.original = ((byte2 & 0x04) >> 2);

				//------------------------------------
				// emphasis
				header.emphasis = (byte2 & 0x03);

				if (header.emphasis == 2)
				{
					goto BAD_HEADER;
				}

				//------------------------------------
				// frame size
				if (header.layer == EX_LAYER_I)
				{
					header.size = ((12000 * header.bitrate / header.samplerate) + header.padding) * 4;
				}
				else
				{
					if ((header.layer == EX_LAYER_III) && (header.mpeg == EX_MPEG_2_5))
					{
						header.size = (72000 * header.bitrate / header.samplerate) + header.padding;
					}
					else
					{
						header.size = (144000 * header.bitrate / header.samplerate) + header.padding;
					}
				}

				//----------------------------------------------
				// solo verifica que dos frames consecutivos son
				// iguales, no que todos los frames sean iguales
				if (bVerifyBitrate == true)
				{
					WORD	next_bitrate;
					
					k += header.size;

					while (k < ml_BufferLen - 1)
					{
						byte1 = *(pBuffer + k);
						byte2 = *(pBuffer + k + 1);

						if ((byte1 == 0xFF) && ((byte2 & 0xE0) == 0xE0))
						{
							//------------------------------------
							// verify MPEG
							if (((byte2 & 0x18) >> 3) == EX_MPEG_FAIL)
							{
								k++;
								continue;
							}

							//------------------------------------
							// verify LAYER
							if (((byte2 & 0x06) >> 1) == EX_LAYER_FAIL)
							{
								k++;
								continue;
							}
							
							byte2 = *(pBuffer + k + 2);

							//------------------------------------
							// bitrate_index
							index = ((byte2 & 0xF0) >> 4);

							if (index == 15)
							{
								k++;
								continue;
							}
							else
							{
								if (header.mpeg == EX_MPEG_1)
								{
									next_bitrate = bitrate_table[header.layer - 1][index];
								}
								else
								{
									next_bitrate = bitrate_table[3 + ((header.layer - 1) >> 1)][index];
								}
							}

							if (next_bitrate == header.bitrate)
							{
								header.variable = 0;
							}
							else
							{
								header.variable = 1;
							}

							return true;
						}
						else
						{
							k++;
						}
					}

					return false;
				}
				else
				{
					header.variable = 0;
					return true;
				}
			}
			else
			{
BAD_HEADER:
				k++;
			}
		}
	}

	return false;
}

bool CExTag::CloseFile()
{
	if (mb_FileOpen)
	{
		if (close(nFile) == 0)
		{
			mb_FileOpen = false;
			return true;
		}
	}

	return false;
}

bool CExTag::FreeBuffer()
{
	if (NULL != pBuffer)
	{
		free (pBuffer);
		pBuffer = NULL;
		ml_BufferLen = 0;
	}

	return true;
}
