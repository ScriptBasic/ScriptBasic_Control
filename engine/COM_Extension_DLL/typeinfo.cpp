#include <windows.h>
#include <stdio.h>
#include <list>
#include <string>
#include <map>

#include <comdef.h> 
#include <AtlBase.h>
#include <AtlConv.h>
#include <atlsafe.h>
 
#include "resource.h" 

#define nullptr 0
#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

//variant types use table: http://msdn.microsoft.com/en-us/library/windows/desktop/ms221170%28v=vs.85%29.aspx
//param flags: http://msdn.microsoft.com/en-us/library/windows/desktop/ms221019.aspx

extern char* __B2C(BSTR bString);

/* BSTR to C String conversion
char* __B2C(BSTR bString)
{
	int i;
	int n = (int)SysStringLen(bString);
	char *sz;
	sz = (char *)malloc(n + 1);

	for(i = 0; i < n; i++){
		sz[i] = (char)bString[i];
	}
	sz[i] = 0;
	return sz;
}*/

// BSTR to std::String conversion
std::string __B2S(BSTR bString)
{
	char *sz = __B2C(bString);
	std::string tmp = sz;
	free(sz);
	return tmp;
}

std::string TypeAsString(VARTYPE vt){
  switch (vt){
		case VT_EMPTY: return "VT_EMPTY";
		case VT_NULL: return "VT_NULL";
		case VT_I2: return "VT_I2";
		case VT_I4: return "VT_I4";
		case VT_R4: return "VT_R4";
		case VT_R8: return "VT_R8";
		case VT_CY: return "VT_CY";
		case VT_DATE: return "VT_DATE";
		case VT_BSTR: return "VT_BSTR";
		case VT_DISPATCH: return "VT_DISPATCH";
		case VT_ERROR: return "VT_ERROR";
		case VT_BOOL: return "VT_BOOL";
		case VT_VARIANT: return "VT_VARIANT";
		case VT_DECIMAL: return "VT_DECIMAL";
		case VT_RECORD: return "VT_RECORD";
		case VT_UNKNOWN: return "VT_UNKNOWN";
		case VT_I1: return "VT_I1";
		case VT_UI1: return "VT_UI1";
		case VT_UI2: return "VT_UI2";
		case VT_UI4: return "VT_UI4";
		case VT_INT: return "VT_INT";
		case VT_UINT: return "VT_UINT";
		case VT_VOID: return "VT_VOID";
		case VT_HRESULT: return "VT_HRESULT";
		case VT_PTR: return "VT_PTR";
		case VT_SAFEARRAY: return "VT_SAFEARRAY";
		case VT_CARRAY: return "VT_CARRAY";
		case VT_USERDEFINED: return "VT_USERDEFINED";
		case VT_LPSTR: return "VT_LPSTR";
		case VT_LPWSTR: return "VT_LPWSTR";
		case VT_BLOB: return "VT_BLOB";
		case VT_STREAM: return "VT_STREAM";
		case VT_STORAGE: return "VT_STORAGE";
		case VT_STREAMED_OBJECT: return "VT_STREAMED_OBJECT";
		case VT_STORED_OBJECT: return "VT_STORED_OBJECT";
		case VT_BLOB_OBJECT: return "VT_BLOB_OBJECT";
		case VT_CF: return "VT_CF"; 
		case VT_CLSID: return "VT_CLSID";
		case VT_VECTOR: return "VT_VECTOR";
		case VT_ARRAY: return "VT_ARRAY";
		default: return "???";
  }
}

std::string TypeAsVBString(VARTYPE vt){
  switch (vt){
		case VT_EMPTY: return "Empty";
		case VT_NULL: return "Null";
		case VT_I2: return "Integer";
		case VT_I4: return "Long";
		case VT_R4: return "Long";
		case VT_R8: return "Currency";
		case VT_CY: return "Currency";
		case VT_DATE: return "Date";
		case VT_BSTR: return "String";
		case VT_DISPATCH: return "Object";
		case VT_ERROR: return "Error";
		case VT_BOOL: return "Boolean";
		case VT_VARIANT: return "Variant";
		case VT_DECIMAL: return "VT_DECIMAL";
		case VT_RECORD: return "VT_RECORD";
		case VT_UNKNOWN: return "VT_UNKNOWN";
		case VT_I1: return "Byte";
		case VT_UI1: return "Byte";
		case VT_UI2: return "Integer";
		case VT_UI4: return "Long";
		case VT_INT: return "Integer";
		case VT_UINT: return "Integer";
		case VT_VOID: return "Void";
		case VT_HRESULT: return "Hresult";
		case VT_PTR: return "Object";
		case VT_SAFEARRAY: return "Array";
		case VT_CARRAY: return "CArray";
		case VT_USERDEFINED: return "VT_USERDEFINED";
		case VT_LPSTR: return "VT_LPSTR";
		case VT_LPWSTR: return "VT_LPWSTR";
		case VT_BLOB: return "VT_BLOB";
		case VT_STREAM: return "VT_STREAM";
		case VT_STORAGE: return "VT_STORAGE";
		case VT_STREAMED_OBJECT: return "VT_STREAMED_OBJECT";
		case VT_STORED_OBJECT: return "VT_STORED_OBJECT";
		case VT_BLOB_OBJECT: return "VT_BLOB_OBJECT";
		case VT_CF: return "VT_CF"; 
		case VT_CLSID: return "VT_CLSID";
		case VT_VECTOR: return "VT_VECTOR";
		case VT_ARRAY: return "VT_ARRAY";
		default: return "???";
  }
}

std::string VariantValueAsString(VARIANT vt){
  
	if( vt.vt == VT_BOOL ){
		if(vt.boolVal == 0) return "False";
		return "True";
	}
	
	/*
	char szBuff[100]={0};
	if(vt.vt == VT_DECIMAL){
		sprintf("%f",szBuff,vt.decVal);
		return szBuff;
	}	
	*/

	try{
		_variant_t vValue = vt;
		_bstr_t bstr = vValue;

		char* c = bstr;
		std::string tmp = c;
		return tmp;
	}catch(...){
		return "Error reading value for type " + TypeAsVBString(vt.vt);
	}

}

int __stdcall ShowVariantType(VARIANT *pVal, char* buf, int bufSz)
{
#pragma EXPORT
	
	std::string retVal;

	char b[20];
	sprintf(b, "0x%x = ", pVal->vt);
	retVal = b;

	VARTYPE vt = pVal->vt;

	if( V_ISARRAY(pVal) ){
		retVal += "VT_ARRAY ";
		vt &= ~VT_ARRAY; //remove the flag
	}

	if( V_ISBYREF(pVal) ){
		retVal += "VT_BYREF ";
		vt &= ~VT_BYREF;  
	}

	retVal += TypeAsString(vt);
	const char* c = retVal.c_str();
	int sz = strlen(c);

	if(sz < bufSz){
		strcpy(buf,c);
	}

	return sz;

}

HRESULT TypeName(IDispatch* pDisp, std::string *retVal)
{
    HRESULT hr = S_OK;
	UINT count = 0;

    CComPtr<IDispatch> spDisp(pDisp);
    if(!spDisp)
        return E_INVALIDARG;

    CComPtr<ITypeInfo> spTypeInfo;
    hr = spDisp->GetTypeInfo(0, 0, &spTypeInfo);

    if(SUCCEEDED(hr) && spTypeInfo)
    {
        CComBSTR funcName;
        hr = spTypeInfo->GetDocumentation(-1, &funcName, nullptr, nullptr, nullptr);
        if(SUCCEEDED(hr) && funcName.Length()> 0 )
        {
          char* c = __B2C(funcName);
		  *retVal = c;
		  free(c);
        }         
    }

    return hr;

}

HRESULT TypeName2(ITypeInfo *spTypeInfo, std::string *retVal)
{
    CComBSTR funcName;
    HRESULT hr = spTypeInfo->GetDocumentation(-1, &funcName, nullptr, nullptr, nullptr);
    if(SUCCEEDED(hr) && funcName.Length()> 0 )
    {
	  *retVal = __B2S(funcName);
    }         

    return hr;
}


//FUNCFLAG_FRESRICTED=1, 
#define isRestricted(x)   (((x->wFuncFlags & 0x01)==0) ? false : true)
#define isHidden(x)       (((x->wFuncFlags & 0x40)==0) ? false : true)
#define isNonBrowsable(x) (((x->wFuncFlags & 0x400)==0) ? false : true)

HRESULT GetIDispatchMethods(IDispatch* pDisp, std::list<std::string> &methods)
{
    HRESULT hr = S_OK;
	UINT count = 0;
	char buf[100];
    std::string tmp;

	USES_CONVERSION;

    CComPtr<IDispatch> spDisp(pDisp);
    if(!spDisp)
        return E_INVALIDARG;

    CComPtr<ITypeInfo> spTypeInfo;
	hr = spDisp->GetTypeInfo(0, 0, &spTypeInfo);

    if(SUCCEEDED(hr) && spTypeInfo)
    {
        TYPEATTR *pTatt = nullptr;
        hr = spTypeInfo->GetTypeAttr(&pTatt);
		
		if(SUCCEEDED(hr) && pTatt)
        {
			/*if(pTatt->wMajorVerNum != 0 && pTatt->wMinorVerNum != 0)
			{
				sprintf(buf, "Version: %d.%d", pTatt->wMajorVerNum, pTatt->wMinorVerNum);
				tmp= buf;
				methods.push_back(tmp);
			}

			WCHAR* pwszProgID = NULL;
			if (!FAILED(hr = (ProgIDFromCLSID(pTatt->guid,&pwszProgID)))) //not right one..
			{
				tmp = W2A(pwszProgID);
				methods.push_back("ProgID: " + tmp);
				CoTaskMemFree(pwszProgID);
			}

			if (!FAILED(hr = (StringFromCLSID(pTatt->guid,&pwszProgID))))
			{
				tmp = W2A(pwszProgID);
				methods.push_back("CLSID: " + tmp);
				CoTaskMemFree(pwszProgID);
			}*/

            FUNCDESC *fd = nullptr;
			
            for(int i = 0; i < pTatt->cFuncs; ++i)
            {
                hr = spTypeInfo->GetFuncDesc(i, &fd);
                if(SUCCEEDED(hr) && fd && !isRestricted(fd) && !isHidden(fd) && !isNonBrowsable(fd)) 
                {
                    CComBSTR funcName;
                    spTypeInfo->GetDocumentation(fd->memid, &funcName, nullptr, nullptr, nullptr);
                    if(funcName.Length()>0)
                    {
						tmp = __B2S(funcName.m_str);
						//if(tmp == "CopyFile") DebugBreak();

						if(fd->invkind == INVOKE_PROPERTYGET) tmp = "Get " + tmp;
						if(fd->invkind == INVOKE_PROPERTYPUT) tmp = "Let " + tmp;
						if(fd->invkind == INVOKE_PROPERTYPUTREF) tmp = "Set " + tmp;

						if(fd->invkind == INVOKE_FUNC){
							if(fd->elemdescFunc.tdesc.vt == VT_VOID) tmp = "Sub " + tmp; 
							 else tmp = "Function " + tmp;		
						}

						if(fd->cParams == 0){
								tmp+= "()";
						}else
						{
							tmp+="(";

							unsigned int nofFuncNames=0;
							int numFuncs = fd->cParams+1;
							BSTR* funcNames = new BSTR[numFuncs];
							
							if (funcNames) {

								hr = spTypeInfo->GetNames(fd->memid, funcNames, numFuncs, &nofFuncNames);
								if(hr==S_OK){
									for(int j =0 ; j <  fd->cParams ; j++){
										PARAMDESC p = fd->lprgelemdescParam[j].paramdesc;
										TYPEDESC t = fd->lprgelemdescParam[j].tdesc;

										std::string argx = TypeAsVBString(t.vt);
										if(t.vt == VT_USERDEFINED && t.hreftype != 0){
											CComPtr<ITypeInfo> ti2;
											hr = spTypeInfo->GetRefTypeInfo(t.hreftype, &ti2);
											if(hr==S_OK) hr = TypeName2(ti2, &argx);
										}

										std::string fns;
										if( (j+1) < nofFuncNames) fns	= __B2S(funcNames[j+1]);
									
										if( (p.wParamFlags & PARAMFLAG_FOPT) != 0) tmp+="["; 
										tmp += fns + (fns.length() == 0 ? "" : " As ") + argx;
									 
										if( (p.wParamFlags & PARAMFLAG_FHASDEFAULT) != 0){
											std::string def = VariantValueAsString(p.pparamdescex->varDefaultValue);
											if(def.length() > 0) tmp+= " = " + def;
										}

										if( (p.wParamFlags & PARAMFLAG_FOPT) != 0) tmp+="]"; 
										
										if(j+1 < fd->cParams) tmp+= ", ";
									}
								}

								for (int i=0 ;i < nofFuncNames; i++) SysFreeString(funcNames[i]);
								delete [] funcNames;
							}

							tmp+=")";
						}

						//add return type..
						TYPEDESC tt = fd->elemdescFunc.tdesc;
						if(tt.vt != VT_VOID){
							std::string tas = TypeAsVBString(tt.vt);
							if(tt.vt == VT_USERDEFINED && tt.hreftype != 0){
								CComPtr<ITypeInfo> ti2;
								hr = spTypeInfo->GetRefTypeInfo(tt.hreftype, &ti2);
								if(hr==S_OK) hr = TypeName2(ti2, &tas);
							}					 
							if(tas.length() > 0) tmp += " As " + tas;
						}
						 
                        methods.push_back(tmp);
                    }

                    spTypeInfo->ReleaseFuncDesc(fd);
                }
            }

            spTypeInfo->ReleaseTypeAttr(pTatt);
        }
    }

    return hr;

}


std::string methodDump;

// This is the message loop
BOOL CALLBACK MyDlgProc(HWND hWnd,UINT uMsg,WPARAM wParam,LPARAM lParam)
{
	HWND hEdit = GetDlgItem(hWnd, IDC_EDIT1);
	switch(uMsg)
	{
		case WM_INITDIALOG :
			SetWindowText(hEdit, methodDump.c_str());
			break;

		case WM_COMMAND:
			switch(LOWORD(wParam))
			{
				case IDOK: //delete item from listbox
					EndDialog(hWnd,0);
					break;
			}
			break;

		case WM_CLOSE:
			EndDialog(hWnd,0);
			break;
	}
	return FALSE;
}

void __stdcall DescribeInterface(IDispatch* pDisp)
{
	std::list<std::string> methods;
	//IDispatch* pDisp = nullptr;
	HRESULT hr;

	//hr = iUnk->QueryInterface(IID_IDispatch, (void **)&pDisp);
    //if (FAILED(hr)) goto error;
	
	TypeName(pDisp,&methodDump);
	methodDump = "Class: " + methodDump + "\r\n";

	GetIDispatchMethods(pDisp, methods);

	typedef std::list<std::string>::iterator it_type;
	for(it_type iterator = methods.begin(); iterator != methods.end(); iterator++) {
		methodDump+=*iterator + "\r\n"; 
	}

	//display results modally
	int ret = DialogBox(GetModuleHandle("COM.dll"), MAKEINTRESOURCE(IDD_DIALOG1),0,(DLGPROC)MyDlgProc); 
	if(ret==-1) MessageBox(0, methodDump.c_str(), "Create Dialog Failed...",0);
	methodDump.empty();

error:
	;

}




/*
CComSafeArray<double> arr(10);
    arr[0] = 2.0;
    arr[1] = 3.0;
    arr[2] = 5.0;

*/