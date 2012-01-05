//*******************************************************************
// ExTag.cpp: implementation of the CExTag class.
//*******************************************************************

#ifndef MPEG_AUDIO_EXTAG_H_
#define MPEG_AUDIO_EXTAG_H_

#include <windows.h>


#if _MSC_VER > 1000
#pragma once
#endif

//-------------------------------------------------------------------
// mpeg
//-------------------------------------------------------------------
#define		EX_MPEG_2_5		0
#define		EX_MPEG_FAIL	1
#define		EX_MPEG_2		2
#define		EX_MPEG_1		3


//-------------------------------------------------------------------
// layer 
//-------------------------------------------------------------------
#define		EX_LAYER_FAIL	0
#define		EX_LAYER_III	1		/* Layer III	*/
#define		EX_LAYER_II		2		/* Layer II		*/
#define		EX_LAYER_I		3		/* Layer I		*/


//-------------------------------------------------------------------
// crc
//-------------------------------------------------------------------
#define		EX_WITH_CRC		0
#define		EX_NO_CRC		1


//-------------------------------------------------------------------
// padding
//-------------------------------------------------------------------
#define		EX_NO_PADDED	0
#define		EX_PADDED		1


//-------------------------------------------------------------------
// channel mode
//-------------------------------------------------------------------
#define		EX_MODE_SINGLE_CHANNEL	3	/* single channel				*/
#define		EX_MODE_DUAL_CHANNEL	2	/* dual channel					*/
#define		EX_MODE_JOINT_STEREO	1	/* joint (MS/intensity) stereo	*/
#define		EX_MODE_STEREO			0	/* normal LR stereo				*/


//-------------------------------------------------------------------
// copyright
//-------------------------------------------------------------------
#define		EX_NO_COPYRIGHT		0
#define		EX_WITH_COPYRIGHT	1


//-------------------------------------------------------------------
// original
//-------------------------------------------------------------------
#define		EX_COPY		0
#define		EX_ORIGINAL	1


//-------------------------------------------------------------------
// emphasis
//-------------------------------------------------------------------
#define		EX_EMPHASIS_NONE		0	/* no emphasis					*/
#define		EX_EMPHASIS_50_15_US	1	/* 50/15 microseconds emphasis	*/
#define		EX_EMPHASIS_CCITT_J_17	3	/* CCITT J.17 emphasis			*/



struct ex_frame_header
{
	BYTE	mpeg;			/* version mpeg						*/
	BYTE	layer;			/* audio layer (1, 2, or 3)			*/
	BYTE	crc;			/* CRC								*/
	WORD	bitrate;		/* stream bitrate (kbps)			*/
	INT		samplerate;		/* sampling frequency (Hz)			*/
	BYTE	padding;		/* Padding bit						*/
	BYTE	mode;			/* channel mode (see above)			*/
	BYTE	copyright;		/* Copyright						*/
	BYTE	original;		/* Original							*/
	BYTE	emphasis;		/* de-emphasis to use (see above)	*/
	INT		size;			/* tamaño							*/
	BYTE	variable;		/* Bitrate variable					*/
};



class CExTag  
{
public:
	bool FreeBuffer();
	
	CExTag();
	~CExTag();

	bool	CloseFile();
	bool	ReadFile();
	bool	GetFrameHeader(bool bVerifyBitrate);	
	bool	SetBufferLegth (long length);
	bool	OpenFile (char *szPath);

	ex_frame_header	header;
	
private:
	char			ms_PathFile [MAX_PATH + 1];
	unsigned char * pBuffer;
	long			ml_BufferLen;
	int				nFile;
	bool			mb_FileOpen;
};

#endif
