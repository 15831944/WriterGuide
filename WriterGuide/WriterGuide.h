
// WriterGuide.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CWriterGuideApp:
// See WriterGuide.cpp for the implementation of this class
//

class CWriterGuideApp : public CWinApp
{
public:
	CWriterGuideApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern CWriterGuideApp theApp;