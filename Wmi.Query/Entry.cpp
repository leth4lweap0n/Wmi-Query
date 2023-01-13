#define _WIN32_DCOM
#include <iostream>
#include <comdef.h>
#include <WbemIdl.h>
#include <vector>
#include <string>

#pragma comment(lib, "wbemuuid.lib")
#define WMI_QUERY(query, prop_name_of_result_object) QueryAndPrintResult(query, prop_name_of_result_object)
enum class e_WmiQueryError {
	None,
	BadQueryFailure,
	PropertyExtractionFailure,
	ComInitializationFailure,
	SecurityInitializationFailure,
	IWbemLocatorFailure,
	IWbemServiceConnectionFailure,
	BlanketProxySetFailure,
};

struct SWmiQueryResult
{
	std::vector<std::string> ResultList;
	e_WmiQueryError Error = e_WmiQueryError::None;
	std::wstring ErrorDescription;
};


SWmiQueryResult GetWmiQueryResult(const std::wstring& wmi_query, const std::wstring& prop_name_of_result_object, bool allow_empty_items = false) {

	SWmiQueryResult retVal;
	retVal.Error = e_WmiQueryError::None;
	retVal.ErrorDescription = L"";
	HRESULT hres;
	IWbemLocator* pLoc = nullptr;
	IWbemServices* pSvc = nullptr;
	IEnumWbemClassObject* pEnumerator = nullptr;
	IWbemClassObject* pcls_obj = nullptr;
	VARIANT vt_prop;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		retVal.Error = e_WmiQueryError::ComInitializationFailure;
		retVal.ErrorDescription = L"Failed to initialize COM library. Error code : " + std::to_wstring(hres);
	}
	else
	{
		hres = CoInitializeSecurity(
			nullptr,
			-1,                         
			nullptr,                    
			nullptr,                    
			RPC_C_AUTHN_LEVEL_DEFAULT,  
			RPC_C_IMP_LEVEL_IMPERSONATE,
			nullptr,                    
			EOAC_NONE,                  
			nullptr                     
		);
		if (FAILED(hres))
		{
			retVal.Error = e_WmiQueryError::SecurityInitializationFailure;
			retVal.ErrorDescription = L"Failed to initialize security. Error code : " + std::to_wstring(hres);
		}
		else
		{
			pLoc = nullptr;
			hres = CoCreateInstance(
				CLSID_WbemLocator,
				nullptr,
				CLSCTX_INPROC_SERVER,
				IID_IWbemLocator, reinterpret_cast<LPVOID*>(&pLoc));
			if (FAILED(hres))
			{
				retVal.Error = e_WmiQueryError::IWbemLocatorFailure;
				retVal.ErrorDescription = L"Failed to create IWbemLocator object. Error code : " + std::to_wstring(hres);
			}
			else
			{
				pSvc = nullptr;
				hres = pLoc->ConnectServer(
					_bstr_t(L"ROOT\\CIMV2"), 
					nullptr,                
					nullptr,                
					nullptr,       
					NULL,                   
					nullptr,     
					nullptr,           
					&pSvc                   
				);

				if (FAILED(hres))
				{
					retVal.Error = e_WmiQueryError::IWbemServiceConnectionFailure;
					retVal.ErrorDescription = L"Could not connect to Wbem service.. Error code : " + std::to_wstring(hres);
				}
				else
				{
					hres = CoSetProxyBlanket(
						pSvc,                        
						RPC_C_AUTHN_WINNT,           
						RPC_C_AUTHZ_NONE,            
						nullptr,                     
						RPC_C_AUTHN_LEVEL_CALL,      
						RPC_C_IMP_LEVEL_IMPERSONATE, 
						nullptr,                     
						EOAC_NONE                    
					);
					if (FAILED(hres))
					{
						retVal.Error = e_WmiQueryError::BlanketProxySetFailure;
						retVal.ErrorDescription = L"Could not set proxy blanket. Error code : " + std::to_wstring(hres);
					}
					else
					{
						pEnumerator = nullptr;
						hres = pSvc->ExecQuery(
							bstr_t("WQL"),
							bstr_t(wmi_query.c_str()),
							WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
							nullptr,
							&pEnumerator);

						if (FAILED(hres))
						{
							retVal.Error = e_WmiQueryError::BadQueryFailure;
							retVal.ErrorDescription = L"Bad query. Error code : " + std::to_wstring(hres);
						}
						else
						{
							pcls_obj = nullptr;
							ULONG u_return = 0;

							while (pEnumerator)
							{
								HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
									&pcls_obj, &u_return);

								if (0 == u_return)
								{
									break;
								}
								hr = pcls_obj->Get(prop_name_of_result_object.c_str(), 0, &vt_prop, nullptr, nullptr);
								if (S_OK != hr) {
									retVal.Error = e_WmiQueryError::PropertyExtractionFailure;
									retVal.ErrorDescription = L"Couldn't extract property: " + prop_name_of_result_object + L" from result of query. Error code : " + std::to_wstring(hr);
								}
								else {
									std::string val = [](const VARIANT& v) {
										switch (v.vt) {
										case VT_BSTR:
											return std::string(_bstr_t(v.bstrVal));
										case VT_I4:
											return std::to_string(v.intVal);
										case VT_UI4:
											return std::to_string(v.uintVal);
										case VT_I8:
											return std::to_string(v.llVal);
										case VT_UI8:
											return std::to_string(v.ullVal);
										case VT_R4:
											return std::to_string(v.fltVal);
										case VT_R8:
											return std::to_string(v.dblVal);
										case VT_BOOL:
											return std::to_string(v.boolVal);
										default:
											return std::string();
										}
									}(vt_prop);

									if (val.empty()) {
										if (allow_empty_items) {
											retVal.ResultList.emplace_back("");
										}
									}
									else {
										retVal.ResultList.emplace_back(val);
									}
								}
							}
						}
					}
				}
			}
		}
	}

	VariantClear(&vt_prop);
	if (pcls_obj)
		pcls_obj->Release();
	if (pSvc)
		pSvc->Release();
	if (pLoc)
		pLoc->Release();
	if (pEnumerator)
		pEnumerator->Release();
	CoUninitialize();
	return retVal;
}

std::vector<std::string> QueryAndPrintResult(const std::wstring& query, const std::wstring& prop_name_of_result_object)
{
	const SWmiQueryResult res = GetWmiQueryResult(query, prop_name_of_result_object);
	if (res.Error != e_WmiQueryError::None) {
		std::wcout << "Got this error while executing query: " << std::endl;
		std::wcout << res.ErrorDescription << std::endl;
		return {};
	}
	return res.ResultList;
}


int main(int argc, char** argv)
{
	/*const auto x = WMI_QUERY(L"SELECT SerialNumber FROM Win32_PhysicalMedia", L"SerialNumber");*/
	const auto y = WMI_QUERY(L"SELECT * FROM Win32_Processor", L"ProcessorId");
	const auto z = WMI_QUERY(L"SELECT * FROM Win32_BaseBoard ", L"SerialNumber");
	std::string hwid;
	for (auto& i : y) 
		hwid += i;
	for (auto& i : z) 
		hwid += i;
	std::cout << hwid << std::endl;
	system("pause");
}

