#include "stdafx.h"
#include "ActiveXCOM.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}

/*****************************************************************************
 *                                                                           *
 * Potencia                                                                  *
 * Retorna o valor de um número elevado a outro.                             *
 *                                                                           *
 *****************************************************************************/
double Potencia( double x, double y )
{
	double nResultado;

	nResultado = exp( (y * log(x)) );

	return ( nResultado );
}

void UpperCase(TCHAR * szA)
{
	size_t nLen = _tcslen(szA);

	for (size_t n = 0; n < nLen; n++)
		szA[n] = toupper(szA[n]);
}

HRESULT __fastcall AnsiToUnicode(LPCSTR pszA, LPOLESTR* ppszW)
{

    ULONG cCharacters;
    DWORD dwError;

    if (NULL == pszA)
    {
        *ppszW = NULL;
        return NOERROR;
    }

    cCharacters =  strlen(pszA)+1;

    *ppszW = (LPOLESTR) CoTaskMemAlloc(cCharacters*2);
    if (NULL == *ppszW)
        return E_OUTOFMEMORY;

    // Covert to Unicode.
    if (0 == MultiByteToWideChar(CP_ACP, 0, pszA, cCharacters, *ppszW, cCharacters))
    {
        dwError = GetLastError();
        CoTaskMemFree(*ppszW);
        *ppszW = NULL;
        return HRESULT_FROM_WIN32(dwError);
    }

    return NOERROR;
}

HRESULT __fastcall UnicodeToAnsi(LPCOLESTR pszW, LPSTR* ppszA)
{

    ULONG cbAnsi, cCharacters;
    DWORD dwError;

    if (pszW == NULL)
    {
        *ppszA = NULL;
        return NOERROR;
    }

    cCharacters = wcslen(pszW)+1;
    cbAnsi = cCharacters*2;

    *ppszA = (LPSTR) CoTaskMemAlloc(cbAnsi);
    if (NULL == *ppszA)
        return E_OUTOFMEMORY;

    if (0 == WideCharToMultiByte(CP_ACP, 0, pszW, cCharacters, *ppszA,
                  cbAnsi, NULL, NULL))
    {
        dwError = GetLastError();
        CoTaskMemFree(*ppszA);
        *ppszA = NULL;
        return HRESULT_FROM_WIN32(dwError);
    }
    return NOERROR;
} 

/*****************************************************************************
 * Classe CConnection                                                        *
 *                                                                           *
 *****************************************************************************/

/*****************************************************************************
 * Construtor                                                                *
 *                                                                           *
 *****************************************************************************/
CActiveXCOM::CActiveXCOM()
{
	m_bCreated = FALSE;
	m_pdisp = NULL;  
	m_punk = NULL; 
	m_hr = NOERROR;
	m_bByRef = FALSE;
}

/*****************************************************************************
 * Destrutor                                                                 *
 *                                                                           *
 *****************************************************************************/
CActiveXCOM::~CActiveXCOM()
{
	if (m_bByRef)
		return;

    if(m_punk) 
	{
		m_punk->Release();
		m_punk = NULL;
	}
    if(m_pdisp) 
	{
		m_pdisp->Release();
		m_pdisp = NULL;
	}
}

/*****************************************************************************
 * CreateObject                                                              *
 * Função que intancia um objeto dinamicamente                               *
 * 
 *****************************************************************************/
UINT CActiveXCOM::CreateObject(TCHAR * pszObjeto)
{
	LPOLESTR pszProgID;
	UINT nError;
    CLSID clsid;                  // CLSID of ActiveX object.    

	#ifndef UNICODE
		AnsiToUnicode(pszObjeto, &pszProgID);
		_tcscpy(m_szCOMName, pszObjeto);
	#else
		pszProgID = pszObjeto; 
		_tcscpy(m_szCOMName, pszObjeto);
	#endif
    
    // Retrieve CLSID from the ProgID that the user specified.
	try
	{
		m_hr = CLSIDFromProgID(pszProgID, &clsid);
		if(FAILED(m_hr))
		{
			nError = E_INVALID_PROGID;
			goto error;
		}

		m_hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER, IID_IUnknown, (void **)&m_punk);
		if(FAILED(m_hr))
		{
			nError = E_INSTANCE_NOT_CREATED;
			goto error;
		}

		// Ask the ActiveX object for the IDispatch interface.
		m_hr = m_punk->QueryInterface(IID_IDispatch, (void **)&m_pdisp);
		if(FAILED(m_hr))
		{
			nError = E_OBJECT_NOT_CREATED;
			goto error;
		}
	}
	catch(...)
	{
		nError = E_OBJECT_NOT_CREATED;
		goto error;
	}

	/*#ifndef UNICODE
		::SysFreeString(pszProgID);
	#endif

    m_punk->Release();*/

	m_punk = NULL;
	m_bCreated = TRUE;
	m_bByRef   = FALSE;

    return S_OK;
    
error:
	m_bCreated = FALSE;

    if(m_punk) 
	{
		m_punk->Release();
		m_punk = NULL;
	}
    if(m_pdisp) 
	{
		m_pdisp->Release();
		m_pdisp = NULL;
	}
	
	#ifndef UNICODE
		#ifndef DEBUG
			::CoTaskMemFree(pszProgID);
		#endif
	#endif

    return nError;
}

void CActiveXCOM::DestroyObject()
{
	m_szCOMName[0] = 0;
	if (m_bByRef)
	{
		m_punk = NULL;
		m_pdisp = NULL;
		return;
	}

	m_bCreated = FALSE;

    if(m_punk) 
	{
		m_punk->Release();
		m_punk = NULL;
	}

    if(m_pdisp) 
	{
		m_pdisp->Release();
		m_pdisp = NULL;
	}
}

UINT CActiveXCOM::ExecuteMethod(TCHAR * pszMetodo, TCHAR * pszArgumentList, ...)
{
	DISPPARAMS dispparams;
	VARTYPE * pVarType;
	EXCEPINFO FAR pExcepInfo;
	DISPID dispid;
	LPOLESTR pOleStr;
	unsigned int uArgErr = 0;

	DWORD dwRetorno = S_OK;

	memset(&dispparams, 0, sizeof(dispparams));

	#ifndef UNICODE
		AnsiToUnicode(pszMetodo, &pOleStr);
	#else
		pOleStr = pszMetodo; 
	#endif

	va_list vl;  
	va_start(vl, pszArgumentList);  
	int niCont = 1;
	
	UINT nLen = (UINT)_tcslen(pszArgumentList);

	for(register UINT i = 0; i < nLen; i++, pszArgumentList++)
	{
		if(*pszArgumentList == ',')
			niCont++;		
	}

	pszArgumentList -= nLen;

	pVarType = new VARTYPE[niCont];

	TCHAR szAux[1024];
	
	memset(szAux, 0, sizeof(szAux));
	int nElement = 0;
	for(int i = 0; i < nLen; i++, pszArgumentList++)
	{

		if(*pszArgumentList != ',')
		{
			_tcscat(szAux, &*pszArgumentList);
		}

		if  (*pszArgumentList == ',' || i + 1 == nLen)
		{
			UpperCase(szAux);

			if (_tcscmp(szAux, _T("int")) == 0 || _tcscmp(szAux, _T("short")) == 0)
				pVarType[nElement] = VT_I2;
			else if (_tcscmp(szAux, _T("long")) == 0)
				pVarType[nElement] = VT_I4;

			else if (_tcscmp(szAux, _T("float")) == 0)
				pVarType[nElement] = VT_R4;

			else if (_tcscmp(szAux, _T("double")) == 0)
				pVarType[nElement] = VT_R8;

			else if (_tcscmp(szAux, _T("bool")) == 0)
				pVarType[nElement] = VT_BOOL;

			else if (_tcscmp(szAux, _T("bstr")) == 0)
				pVarType[nElement] = VT_BSTR;

			else if (_tcscmp(szAux, _T("idispatch")) == 0)
				pVarType[nElement] = VT_DISPATCH;

			else if(_tcscmp(szAux, _T("pint")) == 0)
				pVarType[nElement] = VT_BYREF | VT_I2;

			else if (_tcscmp(szAux, _T("plong")) == 0)
				pVarType[nElement] = VT_BYREF | VT_I4;

			else if (_tcscmp(szAux, _T("pbool")) == 0)
				pVarType[nElement] = VT_BYREF | VT_BOOL;

			else if (_tcscmp(szAux, _T("pbstr")) == 0)
				pVarType[nElement] = VT_BYREF | VT_BSTR;

			else if(_tcscmp(szAux, _T("uint")) == 0)
				pVarType[nElement] = VT_UI2;

			else if (_tcscmp(szAux, _T("ulong")) == 0)
				pVarType[nElement] = VT_UI4;

			else if(_tcscmp(szAux, _T("puint")) == 0)
				pVarType[nElement] = VT_BYREF | VT_UI2;

			else if (_tcscmp(szAux, _T("pulong")) == 0)
				pVarType[nElement] = VT_BYREF | VT_UI4;

			else if (_tcscmp(szAux, _T("pfloat")) == 0)
				pVarType[nElement] =  VT_BYREF | VT_R4;

			else if (_tcscmp(szAux, _T("pdouble")) == 0)
				pVarType[nElement] =  VT_BYREF | VT_R8;

			else if(_tcscmp(szAux, _T("variant")) == 0)
				pVarType[nElement] = VT_BYREF | VT_VARIANT;

			else if (_tcscmp(szAux, _T("pvariant")) == 0)
				pVarType[nElement] = VT_BYREF | VT_VARIANT;

			else if(_tcscmp(szAux, _T("safearray")) == 0)
				pVarType[nElement] = VT_SAFEARRAY; //VT_ARRAY | VT_VARIANT;

			else if (_tcscmp(szAux, _T("psafearray")) == 0)
				pVarType[nElement] = VT_BYREF | VT_SAFEARRAY;

			nElement++;
			memset(szAux, 0, sizeof(szAux));
		}
	}

	dispparams.cArgs = niCont;
	dispparams.rgvarg = new VARIANTARG[niCont];
	dispparams.cNamedArgs = 0;

	int nElementVarType = 0;
	for (nElement = niCont-1; nElement >= 0; nElement--)
	{
		dispparams.rgvarg[nElement].vt = pVarType[nElementVarType];

		switch(pVarType[nElementVarType])
		{
			case VT_I2:
				dispparams.rgvarg[nElement].iVal = va_arg(vl, short);
				break;
			case VT_I4:
				dispparams.rgvarg[nElement].lVal = va_arg(vl, long);
				break;
			case VT_R4:
				dispparams.rgvarg[nElement].fltVal = va_arg(vl, float);
				break;
			case VT_R8:
				dispparams.rgvarg[nElement].dblVal = va_arg(vl, double);
				break;
			case VT_BOOL:
				dispparams.rgvarg[nElement].boolVal = va_arg(vl, short);
				break;
			case VT_BSTR:
				#ifdef UNICODE
					dispparams.rgvarg[nElement].bstrVal = va_arg(vl, TCHAR*);
				#else
					dispparams.rgvarg[nElement].bstrVal = _bstr_t(va_arg(vl, TCHAR*));
				#endif
				break;

			case VT_DISPATCH:
				dispparams.rgvarg[nElement].pdispVal = va_arg(vl, LPDISPATCH);
				break;

			case VT_BYREF | VT_I2:
				dispparams.rgvarg[nElement].piVal = va_arg(vl, short*);
				break;
			case VT_BYREF | VT_I4:
				dispparams.rgvarg[nElement].plVal = va_arg(vl, long*);
				break;
			case VT_BYREF | VT_R4:
				dispparams.rgvarg[nElement].pfltVal = va_arg(vl, float*);
				break;
			case VT_BYREF | VT_R8:
				dispparams.rgvarg[nElement].pdblVal = va_arg(vl, double*);
				break;
			case VT_BYREF | VT_BOOL:
				dispparams.rgvarg[nElement].pboolVal = va_arg(vl, short*);
				break;
			case VT_BYREF | VT_BSTR:
				#ifdef UNICODE
					dispparams.rgvarg[nElement].pbstrVal = va_arg(vl, TCHAR**);
				#else
					//dispparams.rgvarg[nElement].pbstrVal = _bstr_t(va_arg(vl, _bstr_t(TCHAR**)));
				#endif
				break;

			case VT_UI2:
				dispparams.rgvarg[nElement].uiVal = va_arg(vl, unsigned short);
				break;
			case VT_UI4:
				dispparams.rgvarg[nElement].ulVal = va_arg(vl, unsigned long);
				break;

			case VT_BYREF | VT_UI2:
				dispparams.rgvarg[nElement].puiVal = va_arg(vl, unsigned short*);
				break;
			case VT_BYREF | VT_UI4:
				dispparams.rgvarg[nElement].pulVal = va_arg(vl, unsigned long*);
				break;

			case VT_BYREF | VT_VARIANT:
				dispparams.rgvarg[nElement].pvarVal = va_arg(vl, VARIANT*);
				break;

			case VT_SAFEARRAY:
			//case VT_ARRAY | VT_VARIANT:
				dispparams.rgvarg[nElement].parray = va_arg(vl, SAFEARRAY*);
				break;

			case VT_BYREF | VT_SAFEARRAY:
				dispparams.rgvarg[nElement].pparray = va_arg(vl, SAFEARRAY**);
				break;
		}

		nElementVarType++;
	}

	try
	{
		m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
		if (FAILED(m_hr))
		{
			va_end(vl);

			goto Finalizar;

			dwRetorno = E_GETIDOFNAMES_ERROR;
		}

		memset(&pExcepInfo, 0, sizeof(pExcepInfo));
		m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispparams, NULL, &pExcepInfo, &uArgErr);

		m_szErroDescription[0] = 0;
		if (FAILED(m_hr))
		{
			va_end(vl);
			goto Finalizar;
			//m_strErroDescription = pExcepInfo.bstrDescription;
			dwRetorno = E_INVOKE_METHOD_ERROR;
		}

		va_end(vl);

		goto Finalizar;
	}
	catch(...)
	{
		va_end(vl);

		goto Finalizar;
	}
	
Finalizar:
	for (int nCont = 0; nCont < niCont; nCont++)
	{
		if (dispparams.rgvarg[nCont].vt == VT_BSTR)
			::SysFreeString(dispparams.rgvarg[nCont].bstrVal);
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif

	delete[] dispparams.rgvarg;
 	delete[] pVarType;

	return dwRetorno;
}

UINT CActiveXCOM::ExecuteMethodRet(TCHAR * pszMetodo, VARIANT &ret, TCHAR * pszArgumentList, ...)
{
	DISPPARAMS dispparams;
	VARTYPE * pVarType;
	DISPID dispid;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	DWORD dwRetorno = S_OK;

	memset(&ret, 0, sizeof(VARIANT));

	#ifndef UNICODE
		AnsiToUnicode(pszMetodo, &pOleStr);
	#else
		pOleStr = pszMetodo; 
	#endif

	va_list vl;  
	va_start(vl, pszArgumentList);  
	int niCont = 1;
	
	UINT nLen = (UINT)_tcslen(pszArgumentList);

	for(register UINT i = 0; i < nLen; i++, pszArgumentList++)
	{
		if(*pszArgumentList == ',')
			niCont++;		
	}

	pszArgumentList -= nLen;

	pVarType = new VARTYPE[niCont];

	TCHAR szAux[1024];
	
	memset(szAux, 0, sizeof(szAux));
	int nElement = 0;
	for(int i = 0; i < nLen; i++, pszArgumentList++)
	{

		if(*pszArgumentList != _T(','))
		{
			_tcscat(szAux, &*pszArgumentList);
		}

		if  (*pszArgumentList == _T(',') || i + 1 == nLen)
		{
			UpperCase(szAux);

			if (_tcscmp(szAux, _T("int")) == 0 || _tcscmp(szAux, _T("short")) == 0)
				pVarType[nElement] = VT_I2;
			else if (_tcscmp(szAux, _T("long")) == 0)
				pVarType[nElement] = VT_I4;

			else if (_tcscmp(szAux, _T("float")) == 0)
				pVarType[nElement] = VT_R4;

			else if (_tcscmp(szAux, _T("double")) == 0)
				pVarType[nElement] = VT_R8;

			else if (_tcscmp(szAux, _T("bool")) == 0)
				pVarType[nElement] = VT_BOOL;

			else if (_tcscmp(szAux, _T("bstr")) == 0)
				pVarType[nElement] = VT_BSTR;

			else if (_tcscmp(szAux, _T("idispatch")) == 0)
				pVarType[nElement] = VT_DISPATCH;

			else if (_tcscmp(szAux, _T("pint")) == 0)
				pVarType[nElement] = VT_BYREF | VT_I2;

			else if (_tcscmp(szAux, _T("plong")) == 0)
				pVarType[nElement] = VT_BYREF | VT_I4;

			else if (_tcscmp(szAux, _T("pbool")) == 0)
				pVarType[nElement] = VT_BYREF | VT_BOOL;

			else if (_tcscmp(szAux, _T("pbstr")) == 0)
				pVarType[nElement] = VT_BYREF | VT_BSTR;

			else if (_tcscmp(szAux, _T("uint")) == 0)
				pVarType[nElement] = VT_UI2;

			else if (_tcscmp(szAux, _T("ulong")) == 0)
				pVarType[nElement] = VT_UI4;

			else if (_tcscmp(szAux, _T("puint")) == 0)
				pVarType[nElement] = VT_BYREF | VT_UI2;

			else if (_tcscmp(szAux, _T("pulong")) == 0)
				pVarType[nElement] = VT_BYREF | VT_UI4;

			else if (_tcscmp(szAux, _T("pfloat")) == 0)
				pVarType[nElement] =  VT_BYREF | VT_R4;

			else if (_tcscmp(szAux, _T("pdouble")) == 0)
				pVarType[nElement] =  VT_BYREF | VT_R8;

			else if (_tcscmp(szAux, _T("variant")) == 0)
				pVarType[nElement] = VT_BYREF | VT_VARIANT;

			else if (_tcscmp(szAux, _T("pvariant")) == 0)
				pVarType[nElement] = VT_BYREF | VT_VARIANT;

			else if (_tcscmp(szAux, _T("safearray")) == 0)
				pVarType[nElement] = VT_SAFEARRAY; //VT_ARRAY | VT_VARIANT;

			else if (_tcscmp(szAux, _T("psafearray")) == 0)
				pVarType[nElement] = VT_BYREF | VT_SAFEARRAY;

			nElement++;
			memset(szAux, 0, sizeof(szAux));
		}
	}

	dispparams.cArgs = niCont;
	dispparams.rgvarg = new VARIANTARG[niCont];
	dispparams.cNamedArgs = 0;

	int nElementVarType = 0;
	for (nElement = niCont-1; nElement >= 0; nElement--)
	{
		dispparams.rgvarg[nElement].vt = pVarType[nElementVarType];

		switch(pVarType[nElementVarType])
		{
			case VT_I2:
				dispparams.rgvarg[nElement].iVal = va_arg(vl, short);
				break;
			case VT_I4:
				dispparams.rgvarg[nElement].lVal = va_arg(vl, long);
				break;
			case VT_R4:
				dispparams.rgvarg[nElement].fltVal = va_arg(vl, float);
				break;
			case VT_R8:
				dispparams.rgvarg[nElement].dblVal = va_arg(vl, double);
				break;
			case VT_BOOL:
				dispparams.rgvarg[nElement].boolVal = va_arg(vl, short);
				break;
			case VT_BSTR:
				#ifdef UNICODE
					dispparams.rgvarg[nElement].bstrVal = va_arg(vl, TCHAR*);
				#else
					dispparams.rgvarg[nElement].bstrVal = _bstr_t(va_arg(vl, TCHAR*));
				#endif
				break;

			case VT_DISPATCH:
				dispparams.rgvarg[nElement].pdispVal = va_arg(vl, LPDISPATCH);
				break;

			case VT_BYREF | VT_I2:
				dispparams.rgvarg[nElement].piVal = va_arg(vl, short*);
				break;
			case VT_BYREF | VT_I4:
				dispparams.rgvarg[nElement].plVal = va_arg(vl, long*);
				break;
			case VT_BYREF | VT_R4:
				dispparams.rgvarg[nElement].pfltVal = va_arg(vl, float*);
				break;
			case VT_BYREF | VT_R8:
				dispparams.rgvarg[nElement].pdblVal = va_arg(vl, double*);
				break;
			case VT_BYREF | VT_BOOL:
				dispparams.rgvarg[nElement].pboolVal = va_arg(vl, short*);
				break;
			case VT_BYREF | VT_BSTR:
				#ifdef UNICODE
					dispparams.rgvarg[nElement].pbstrVal = va_arg(vl, TCHAR**);
				#else
					//dispparams.rgvarg[nElement].pbstrVal = _bstr_t(va_arg(vl, _bstr_t(TCHAR**)));
				#endif
				break;

			case VT_UI2:
				dispparams.rgvarg[nElement].uiVal = va_arg(vl, unsigned short);
				break;
			case VT_UI4:
				dispparams.rgvarg[nElement].ulVal = va_arg(vl, unsigned long);
				break;

			case VT_BYREF | VT_UI2:
				dispparams.rgvarg[nElement].puiVal = va_arg(vl, unsigned short*);
				break;
			case VT_BYREF | VT_UI4:
				dispparams.rgvarg[nElement].pulVal = va_arg(vl, unsigned long*);
				break;

			case VT_BYREF | VT_VARIANT:
				dispparams.rgvarg[nElement].pvarVal = va_arg(vl, VARIANT*);
				break;

			case VT_ARRAY|VT_VARIANT:
				dispparams.rgvarg[nElement].parray = va_arg(vl, SAFEARRAY*);
				break;

			case VT_BYREF | VT_SAFEARRAY:
				dispparams.rgvarg[nElement].pparray = va_arg(vl, SAFEARRAY**);
				break;
		}

		nElementVarType++;
	}

	try
	{
		m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
		if (FAILED(m_hr))
		{
			va_end(vl);

			goto Finalizar;

			dwRetorno = E_GETIDOFNAMES_ERROR;
		}

		m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispparams, &ret, &pExcepInfo, NULL);
		m_szErroDescription[0] = 0;
		if (FAILED(m_hr))
		{
			va_end(vl);

			goto Finalizar;
			//m_strErroDescription = pExcepInfo.bstrDescription;
			dwRetorno = E_INVOKE_METHOD_ERROR;
		}

		va_end(vl);

		goto Finalizar;
	}
	catch(...)
	{
		va_end(vl);

		goto Finalizar;
	}
	
Finalizar:
	for (int nCont = 0; nCont < niCont; nCont++)
	{
		if (dispparams.rgvarg[nCont].vt == VT_BSTR)
			::SysFreeString(dispparams.rgvarg[nCont].bstrVal);
	}
	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;
 	delete[] pVarType;

	return dwRetorno;
}

/*****************************************************************************
 * ExecuteMethodSafe                                                         *
 * Função que executa um método                                              *
 * Esta função recebe um formatador de parâmetros, futuramente será          *
 * substituida por ExecuteMethod, que não precisará do formatador            *
 *****************************************************************************/
UINT CActiveXCOM::ExecuteMethodSafe(TCHAR *pszMetodo, SAFEARRAY *pSA)
{
	DISPPARAMS dispparams;
	DISPID dispid;
	LPOLESTR pOleStr;
	VARIANT vtValor;
	long    nIndices[1];
	int nElement = 0;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszMetodo, &pOleStr);
	#else
		pOleStr = pszMetodo; 
	#endif

	long nLBound, nUBound;
	int niCont;
	
	SafeArrayGetLBound(pSA, 1, &nLBound);
	SafeArrayGetUBound(pSA, 1, &nUBound);
	
	niCont = nUBound - nLBound +1;

	dispparams.cArgs = niCont;
	dispparams.rgvarg = new VARIANTARG[niCont];
	dispparams.cNamedArgs = 0;

	int nElementVarType = 0; 
	VariantInit(&vtValor);
	for (nElement = niCont-1; nElement >= 0; nElement--)
	{
		VariantClear(&vtValor);
		nIndices[0] = nElementVarType;

		SafeArrayGetElement(pSA, nIndices, &vtValor);
		dispparams.rgvarg[nElement] = vtValor;

		nElementVarType++;
	}

	try
	{
		m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
		if (FAILED(m_hr))
		{
			#ifndef UNICODE
				::SysFreeString(pOleStr);
			#endif
			delete[] dispparams.rgvarg;
			return E_GETIDOFNAMES_ERROR;
		}

		m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispparams, NULL, &pExcepInfo, NULL);
		m_szErroDescription[0] = 0;
		if (FAILED(m_hr))
		{
			//m_strErroDescription = pExcepInfo.bstrDescription;
			#ifndef UNICODE
				::SysFreeString(pOleStr);
			#endif
			delete[] dispparams.rgvarg;
			return E_INVOKE_METHOD_ERROR;
		}

		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
	}
	catch(...)
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
	}
	

	return 0;
}

/*****************************************************************************
 * ExecuteMethodNoArg                                                        *
 * Função que executa um método                                              *
 *                                                                           *
 *****************************************************************************/
UINT CActiveXCOM::ExecuteMethodNoArg(TCHAR * pszMetodo)
{
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	LPOLESTR pOleStr;
	EXCEPINFO pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszMetodo, &pOleStr);
	#else
		pOleStr = pszMetodo; 
	#endif

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
		return E_GETIDOFNAMES_ERROR;

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispparamsNoArgs, NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_METHOD_ERROR;
	}
	
	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::ExecuteMethodNoArgRet(TCHAR * pszMetodo, VARIANT &ret)
{
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	LPOLESTR pOleStr;
	EXCEPINFO pExcepInfo;

	memset(&ret, 0, sizeof(VARIANT));

	#ifndef UNICODE
		AnsiToUnicode(pszMetodo, &pOleStr);
	#else
		pOleStr = pszMetodo; 
	#endif

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispparamsNoArgs, &ret, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_METHOD_ERROR;
	}
	
	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif

	return S_OK;
}

/*****************************************************************************
 * SetProperty                                                               *
 * Atribui um valor para uma propriedade                                     *
 *                                                                           *
 *****************************************************************************/
UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, short nsParametro )
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_I2;
	dispparams.rgvarg[0].iVal    = nsParametro;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, float nsParametro)
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_R4;
	dispparams.rgvarg[0].fltVal  = nsParametro;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, DWORD nParametro)
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_UI4;
	dispparams.rgvarg[0].ulVal   = nParametro;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, long  nlParametro )
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_I4;
	dispparams.rgvarg[0].lVal    = nlParametro;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, CURRENCY ncParametro )
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_CY;
	dispparams.rgvarg[0].cyVal.int64   = ncParametro.int64;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, DATE ndParametro)
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_DATE;
	dispparams.rgvarg[0].date    = ndParametro;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, TCHAR * pszParametro)
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	LPOLESTR pOleProperty;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
		AnsiToUnicode(pszParametro, &pOleProperty);
	#else
		pOleStr      = pszProperty; 
		pOleProperty = pszParametro;
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_BSTR;
	dispparams.rgvarg[0].bstrVal = pOleProperty;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		::SysFreeString(dispparams.rgvarg[0].bstrVal);
		#ifndef UNICODE
			::SysFreeString(pOleStr);
			::SysFreeString(pOleProperty);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		::SysFreeString(dispparams.rgvarg[0].bstrVal);
		#ifndef UNICODE
			::SysFreeString(pOleStr);
			::SysFreeString(pOleProperty);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	::SysFreeString(dispparams.rgvarg[0].bstrVal);
	#ifndef UNICODE
		::SysFreeString(pOleStr);
		::SysFreeString(pOleProperty);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, BOOL bParametro)
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt      = VT_BOOL;
	dispparams.rgvarg[0].boolVal = bParametro;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, LPDISPATCH pdisp)
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.rgvarg[0].vt		  = VT_DISPATCH;
	dispparams.rgvarg[0].pdispVal = pdisp;
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUTREF, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::SetProperty(TCHAR * pszProperty, VARIANT vtParametro)
{	
	DISPPARAMS dispparams;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	//Configurando propriedade ConnectionString
	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	memcpy(dispparams.rgvarg, &vtParametro, sizeof(VARIANT));
	dispparams.rgdispidNamedArgs = &mydispid;

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUTREF, &dispparams,NULL, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/*****************************************************************************
 * GetProperty                                                               *
 * Recebe o valor de uma propriedade	                                     *
 *                                                                           *
 *****************************************************************************/
UINT CActiveXCOM::GetProperty(TCHAR* pszProperty, short& pParametro )
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&pVarResult, 0, sizeof(VARIANT));

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	switch(pVarResult.vt)
	{
		case VT_I2:
			pParametro =  pVarResult.iVal;
			break;
		case VT_I4:
			pParametro =  (short)pVarResult.lVal;
			break;
		case VT_UI2:
			pParametro =  (short)pVarResult.uiVal;
			break;
		case VT_UI4:
			pParametro =  (short)pVarResult.ulVal;
			break;
		case VT_R4:
			pParametro =  (short)pVarResult.fltVal;
			break;
		case VT_R8:
			pParametro =  (short)pVarResult.dblVal;
			break;
		case VT_I1:
			pParametro =  (short)pVarResult.intVal;
			break;
		case VT_DECIMAL:
			pParametro = (short)pVarResult.lVal;
			break;
		case VT_NULL:
			pParametro = 0;
			break;
		case VT_EMPTY:
			pParametro = 0;
			break;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR* pszProperty, float& pParametro )
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&pVarResult, 0, sizeof(VARIANT));

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	switch(pVarResult.vt)
	{
		case VT_I2:
			pParametro =  pVarResult.iVal;
			break;
		case VT_I4:
			pParametro =  (float)pVarResult.lVal;
			break;
		case VT_UI2:
			pParametro =  pVarResult.uiVal;
			break;
		case VT_UI4:
			pParametro =  (float)pVarResult.ulVal;
			break;
		case VT_R4:
			pParametro =  pVarResult.fltVal;
			break;
		case VT_R8:
			pParametro =  (float)pVarResult.dblVal;
			break;
		case VT_I1:
			pParametro =  (float)pVarResult.intVal;
			break;
		case VT_DECIMAL:
			if (pVarResult.wReserved1 > 0)
			{
				double nValor, nDecimal;
				nValor    = pVarResult.lVal;
				nDecimal  = Potencia(10, pVarResult.wReserved1);
				nValor = (nValor / nDecimal);
				pParametro = (float)nValor;
			}
			else
				pParametro = (float)pVarResult.lVal;
			break;
		case VT_NULL:
		case VT_EMPTY:
			pParametro = 0;
			break;

	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR* pszProperty, DWORD& pParametro )
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&pVarResult, 0, sizeof(VARIANT));

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	//m_strErroDescription = _T("");
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	switch(pVarResult.vt)
	{
		case VT_I2:
			pParametro =  pVarResult.iVal;
			break;
		case VT_I4:
			pParametro =  pVarResult.lVal;
			break;
		case VT_UI2:
			pParametro =  pVarResult.uiVal;
			break;
		case VT_UI4:
			pParametro =  pVarResult.ulVal;
			break;
		case VT_R4:
			pParametro =  (DWORD)pVarResult.fltVal;
			break;
		case VT_R8:
			pParametro =  (DWORD)pVarResult.dblVal;
			break;
		case VT_I1:
			pParametro =  pVarResult.intVal;
			break;
		case VT_DECIMAL:
			if (pVarResult.wReserved1 > 0)
				pParametro = (long)(pVarResult.lVal / Potencia(10, pVarResult.wReserved1));
			else
				pParametro = pVarResult.lVal;
			break;
		case VT_NULL:
		case VT_EMPTY:
			pParametro = 0;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR * pszProperty, long&  pParametro )
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&pVarResult, 0, sizeof(VARIANT));

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	switch(pVarResult.vt)
	{
		case VT_I2:
			pParametro =  pVarResult.iVal;
			break;
		case VT_I4:
			pParametro =  pVarResult.lVal;
			break;
		case VT_UI2:
			pParametro =  pVarResult.uiVal;
			break;
		case VT_UI4:
			pParametro =  pVarResult.ulVal;
			break;
		case VT_R4:
			pParametro =  (long)pVarResult.fltVal;
			break;
		case VT_R8:
			pParametro =  (long)pVarResult.dblVal;
			break;
		case VT_I1:
			pParametro =  pVarResult.intVal;
			break;
		case VT_DECIMAL:
			pParametro = pVarResult.lVal;
			break;
		case VT_NULL:
		case VT_EMPTY:
			pParametro = 0;
	}

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR * pszProperty, CURRENCY& pParametro )
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&pVarResult, 0, sizeof(VARIANT));

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	pParametro = pVarResult.cyVal;

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR * pszProperty, DATE& pParametro)
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&pVarResult, 0, sizeof(VARIANT));

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	switch(pVarResult.vt)
	{
		case VT_I2:
			pParametro =  pVarResult.iVal;
			break;
		case VT_I4:
			pParametro =  pVarResult.lVal;
			break;
		case VT_UI2:
			pParametro =  pVarResult.uiVal;
			break;
		case VT_UI4:
			pParametro =  pVarResult.ulVal;
			break;
		case VT_R4:
			pParametro =  pVarResult.fltVal;
			break;
		case VT_R8:
			pParametro =  pVarResult.dblVal;
			break;
		case VT_I1:
			pParametro =  pVarResult.intVal;
			break;
		case VT_DECIMAL:
			if (pVarResult.wReserved1 > 0)
				pParametro = pVarResult.lVal / Potencia(10, pVarResult.wReserved1);
			else
				pParametro = pVarResult.lVal;
			break;
		case VT_DATE:
			pParametro =  pVarResult.date;
			break;
		case VT_NULL:
		case VT_EMPTY:
			pParametro = 0;
	}

	
	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR* pszProperty, TCHAR * szParametro)
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&pVarResult, 0, sizeof(VARIANT));

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	if (pVarResult.vt == VT_BSTR)
	{
		UnicodeToAnsi(pVarResult.bstrVal, (LPSTR*)szParametro);
		::SysFreeString(pVarResult.bstrVal);
	}
	else if (pVarResult.vt == VT_BOOL)
	{
		if (pVarResult.boolVal)
		{
			_tcscpy(szParametro, _T("Verdadeiro"));
		}
		else
		{
			_tcscpy(szParametro, _T("Falso"));
		}
	}
	else if (pVarResult.vt == VT_NULL)
		szParametro[0] = 0;

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetPropertyBool(TCHAR * pszProperty, BOOL&  pParametro )
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	pParametro = pVarResult.boolVal;

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR * pszProperty, CActiveXCOM& pActiveX)
{	
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		//m_strErroDescription = pExcepInfo.bstrDescription;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}


	pActiveX.m_bByRef = TRUE;
	pActiveX.m_pdisp =  pVarResult.pdispVal;

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR * pszProperty, long nIndex, CActiveXCOM& pActiveX)
{	
	DISPPARAMS dispparams, dispparamsNoArgs = {NULL, NULL, 0, 0};;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.cNamedArgs = 0;
	dispparams.rgvarg[0].vt = VT_I4;
	dispparams.rgvarg[0].lVal = nIndex;
	dispparams.rgdispidNamedArgs = &mydispid;
	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD , &dispparams, &pVarResult, &pExcepInfo, NULL);
	m_szErroDescription[0] = 0;
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		delete[] dispparams.rgvarg;
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	pActiveX.m_bByRef = TRUE;
	pActiveX.m_pdisp =  pVarResult.pdispVal;

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR * pszProperty, TCHAR * pszField, CActiveXCOM& pActiveX)
{	
	DISPPARAMS dispparams, dispparamsNoArgs = {NULL, NULL, 0, 0};;
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr, pOleField;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
		AnsiToUnicode(pszField, &pOleField);
	#else
		pOleStr   = pszProperty; 
		pOleField = pszField;
	#endif

	memset(&dispparams, 0, sizeof dispparams);
	dispparams.cArgs = 1;
	dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
	memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);
	dispparams.cNamedArgs = 0;
	dispparams.rgvarg[0].vt       = VT_BSTR;
	dispparams.rgvarg[0].bstrVal = pOleField;
	dispparams.rgdispidNamedArgs = &mydispid;
	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		::SysFreeString(dispparams.rgvarg[0].bstrVal);
		#ifndef UNICODE
			::SysFreeString(pOleStr);
			::SysFreeString(pOleField);
		#endif
		delete[] dispparams.rgvarg;
		return E_GETIDOFNAMES_ERROR;
	}

	try
	{
		m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispparams, &pVarResult, &pExcepInfo, NULL);
		if (FAILED(m_hr))
		{
			::SysFreeString(dispparams.rgvarg[0].bstrVal);
			#ifndef UNICODE
				::SysFreeString(pOleStr);
				::SysFreeString(pOleField);
			#endif
			delete[] dispparams.rgvarg;
			return E_INVOKE_PROPERTYPUT_ERROR;
		}
	}
	catch(...)
	{
		int a = 0;
	}

	pActiveX.m_bByRef = TRUE;
	pActiveX.m_pdisp =  pVarResult.pdispVal;

	::SysFreeString(dispparams.rgvarg[0].bstrVal);
	#ifndef UNICODE
		::SysFreeString(pOleStr);
		::SysFreeString(pOleField);
	#endif
	delete[] dispparams.rgvarg;

	return S_OK;
}

UINT CActiveXCOM::GetProperty(TCHAR * pszProperty, VARIANT& vtParametro)
{
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	DISPID dispid;
	DISPID mydispid = DISPID_PROPERTYPUT;
	LPOLESTR pOleStr;
	VARIANT pVarResult;
	EXCEPINFO FAR pExcepInfo;

	#ifndef UNICODE
		AnsiToUnicode(pszProperty, &pOleStr);
	#else
		pOleStr = pszProperty; 
	#endif

	m_hr = m_pdisp->GetIDsOfNames(IID_NULL, &pOleStr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_GETIDOFNAMES_ERROR;
	}

	m_hr = m_pdisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, &pExcepInfo, NULL);
	if (FAILED(m_hr))
	{
		#ifndef UNICODE
			::SysFreeString(pOleStr);
		#endif
		return E_INVOKE_PROPERTYPUT_ERROR;
	}

	vtParametro = pVarResult;

	#ifndef UNICODE
		::SysFreeString(pOleStr);
	#endif
	return S_OK;
}

/*****************************************************************************
 * SetDispatch                                                               *
 * Para que o objeto possa ser configurado por referência                    *
 *                                                                           *
 *****************************************************************************/
void CActiveXCOM::SetDispatch(LPDISPATCH pDispatch)
{
	m_pdisp = pDispatch;
}

TCHAR* CActiveXCOM::GetErrorDescription()
{
	return m_szErroDescription;
}


