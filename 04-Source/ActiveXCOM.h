
#ifdef ACTIVEXCOM_EXPORTS
	#define ACTIVEXCOM_API __declspec(dllexport)
#else
	#define ACTIVEXCOM_API __declspec(dllimport)
#endif

#define E_INVALID_NUMBER_OF_ARGUMENTS	0x00000FA0
#define E_OBJECT_NOT_CREATED			0x00000FA1
#define E_INVALID_PROGID				0x00000FA2
#define E_INSTANCE_NOT_CREATED			0x00000FA3
#define E_GETIDOFNAMES_ERROR			0x00000FA4
#define E_INVOKE_METHOD_ERROR			0x00000FA5
#define E_INVOKE_PROPERTYPUT_ERROR		0x00000FA6

class ACTIVEXCOM_API CActiveXCOM  
{
public:	
	/* Função utilizada para instanciar dinamicamente um método */
	UINT CreateObject(TCHAR * pszObjeto);

	/* Destroy o objecto criado */
	void DestroyObject();	

	/* Executa os métodos do objecto instanciado */
	//UINT ExecuteMethod(TCHAR * pszMetodo, TCHAR * pszTipo, TCHAR * pszParametro, VARIANT &ret); // Executa um método sem parâmetros
	UINT ExecuteMethodNoArg(TCHAR * pszMetodo); // Executa um método sem parâmetros
	UINT ExecuteMethodNoArgRet(TCHAR * pszMetodo, VARIANT &ret); // Executa um método sem parâmetros
	UINT ExecuteMethod(TCHAR * pszMetodo, TCHAR * pszArgumentList, ...);
	UINT ExecuteMethodRet(TCHAR * pszMetodo, VARIANT &ret, TCHAR * pszArgumentList, ...);
	UINT ExecuteMethodSafe(TCHAR * pszMetodo, SAFEARRAY FAR * pSA);


	/* Atribui um valor a uma propriedade */
	UINT SetProperty(TCHAR * pszProperty, short nsParametro );
	UINT SetProperty(TCHAR * pszProperty, long  nlParametro );
	UINT SetProperty(TCHAR * pszProperty, DWORD	  nParametro);
	UINT SetProperty(TCHAR * pszProperty, float  nlParametro );
	UINT SetProperty(TCHAR * pszProperty, CURRENCY ncParametro );
	UINT SetProperty(TCHAR * pszProperty, DATE ndParametro);
	UINT SetProperty(TCHAR * pszProperty, BOOL bParametro);
	UINT SetProperty(TCHAR * pszProperty, TCHAR * pszParametro);
	UINT SetProperty(TCHAR * pszProperty, LPDISPATCH pdisp);
	UINT SetProperty(TCHAR * pszProperty, VARIANT vtParametro);

	/* Recebe o valor de uma propriedade */
	UINT GetProperty(TCHAR* pszProperty, short&			pParametro);
	UINT GetProperty(TCHAR* pszProperty, long&			pParametro);
	UINT GetProperty(TCHAR* pszProperty, DWORD&			pParametro);
	UINT GetProperty(TCHAR* pszProperty, float&			pParametro);
	UINT GetProperty(TCHAR* pszProperty, CURRENCY&		pParametro);
	UINT GetProperty(TCHAR* pszProperty, DATE&			pParametro);
	UINT GetProperty(TCHAR* pszProperty, TCHAR *		szParametro);	
	UINT GetPropertyBool(TCHAR * pszProperty, BOOL&		pParametro);
	UINT GetProperty(TCHAR * pszProperty, CActiveXCOM& pActiveX);
	UINT GetProperty(TCHAR * pszProperty, VARIANT& vtParametro);

	UINT GetProperty(TCHAR * pszProperty, long nIndex, CActiveXCOM& pActiveX);
	UINT GetProperty(TCHAR * pszProperty, TCHAR * pszField, CActiveXCOM& pActiveX);

	/* Retorna o HRESULT da última operação */
	inline HRESULT GetHRESULT() { return m_hr; }

	inline LPDISPATCH GetIDispatch() { return m_pdisp; }
	void SetDispatch(LPDISPATCH pDispatch);

	inline BOOL IsCreated() { return m_bCreated; }

	/* */

public:	
	TCHAR * GetErrorDescription();
	CActiveXCOM();
	virtual ~CActiveXCOM();	

private:
	TCHAR m_szErroDescription[1024];
	TCHAR   m_szCOMName[500];
	HRESULT m_hr;            
	BOOL m_bCreated;				// Variável que controla se o objeto foi ou não criado
	LPDISPATCH m_pdisp;				// Ponteiro para a variável IDispatch.
	LPUNKNOWN m_punk;				// Ponteiro para a variável IUnknown.    
	VARIANT FAR* m_pVarResult;		// Retorna um resultado
	BOOL m_bByRef;
};

