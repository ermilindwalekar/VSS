#include "stdafx.h"

void ReleaseInterface (IUnknown* unkn);


int _tmain(int argc, _TCHAR* argv[])
{
	int returnValue = 1;

	if (argc != 2)
	{
		_tprintf(_T("Usage: simplesnapshot.exe volume\n"));
		_tprintf(_T("example: simplesnapshot.exe [d:\\] or [path_to_mounted_folder]\n"));
		return 1;
	}
	TCHAR volumeName[MAX_PATH] = {0};

	_tprintf(_T("create snapshot for volume %s :\n"), argv[1] );
	_tcscpy_s(volumeName, argv[1] );

	/*
	  OSVERSIONINFO osvi;
	  BOOL bIsWindowsVistaOr7;

	  ZeroMemory(&osvi, sizeof(OSVERSIONINFO));
	  osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	  GetVersionEx(&osvi);
	  bIsWindowsVistaOr7 =  ( (osvi.dwMajorVersion > 6) || ( (osvi.dwMajorVersion == 6) && (osvi.dwMinorVersion >= 0) ));
	*/

#if (_WIN32_WINNT >= _WIN32_WINNT_VISTA)
	/* A program using VSS must run in elevated mode */
	HANDLE hToken;
	OpenProcessToken(GetCurrentProcess(), TOKEN_READ, &hToken);
	DWORD infoLen;

	TOKEN_ELEVATION elevation;
	GetTokenInformation(hToken, TokenElevation, &elevation, sizeof(elevation), &infoLen);
	if (!elevation.TokenIsElevated)
	{
		_tprintf (_T("this program must run in elevated mode\n"));
		return 3;
	}


#else
#error you are using an old version of sdk or not supported operating system
#endif



	if (CoInitialize(NULL) != S_OK)
	{
		_tprintf(_T("CoInitialize failed!\n"));
		return 1;
	}


	// declare all the interfaces used in this program.

	IVssBackupComponents* pBackup = NULL;
	IVssAsync* pAsync             = NULL;
	IVssAsync* pPrepare           = NULL;

	// Create the IVssBackupComponents Interface
	HRESULT result                = CreateVssBackupComponents(&pBackup);


	if (result == S_OK)
	{
		// Initialize for backup
		result = pBackup->InitializeForBackup();

		if (result == S_OK)
		{
			// set the context
			result = pBackup->SetContext(VSS_CTX_BACKUP | VSS_CTX_CLIENT_ACCESSIBLE_WRITERS | VSS_CTX_APP_ROLLBACK);

			if (result != S_OK)
			{
				_tprintf(_T("Error: HRESULT = 0x%08lx\n"), result);
				ReleaseInterface(pBackup);
				return -19;
			}

			// Prompts each writer to send the metadata they have collected
			result = pBackup->GatherWriterMetadata(&pAsync);
			if (result == S_OK)
			{
				_tprintf(_T("Gathering metadata from writers...\n"));
				result = pAsync->Wait();
				if (result == S_OK)
				{
					//
					// Creates a new, empty shadow copy set
					//
					_tprintf(_T("calling StartSnapshotSet...\n"));
					VSS_ID snapshotSetId;
					result = pBackup->StartSnapshotSet(&snapshotSetId);
					if (result != S_OK)
					{
						_tprintf(_T("- Returned HRESULT = 0x%08lx\n"), result);
						ReleaseInterface(pBackup);
					}
					if (result == S_OK)
					{
						_tprintf(_T("AddToSnapshotSet...\n"));
						result = pBackup->AddToSnapshotSet(volumeName, GUID_NULL, &snapshotSetId);
						if (result != S_OK)
						{
							_tprintf(_T("- Returned HRESULT = 0x%08lx\n"), result);
							ReleaseInterface(pBackup);
						}
						if (result == S_OK)
						{
							//
							// Configure the backup operation for Copy with no backup history
							//
							result = pBackup->SetBackupState(false, false, VSS_BT_COPY);
							if (result != S_OK)
							{
								_tprintf(_T("- Returned HRESULT = 0x%08lx\n"), result);
								ReleaseInterface(pBackup);
							}

							if (result == S_OK)
							{
								//
								// Make VSS generate a PrepareForBackup event
								//
								result = pBackup->PrepareForBackup(&pPrepare);
								if (result == S_OK)
								{
									_tprintf(_T("Preparing for backup...\n"));
									result = pPrepare->Wait();
									if (result == S_OK)
									{
										//
										// Commit all snapshots in this set
										//
										IVssAsync* pDoShadowCopy = NULL;
										result                   = pBackup->DoSnapshotSet(&pDoShadowCopy);
										if (result == S_OK)
										{
											_tprintf(_T("Taking snapshots...\n"));
											result = pDoShadowCopy->Wait();
											if (result == S_OK)
											{
												//
												// Get the snapshot device object from the properties
												//
												_tprintf(_T("Get the snapshot device object from the properties...\n"));
												VSS_SNAPSHOT_PROP snapshotProp = {0};
												result = pBackup->GetSnapshotProperties(snapshotSetId, &snapshotProp);
												if (result == S_OK)
												{

													//display out snapsot properties
													_tprintf (_T(" Snapshot Id :")  WSTR_GUID_FMT _T("\n"), GUID_PRINTF_ARG( snapshotProp.m_SnapshotId));
													_tprintf (_T(" Snapshot Set Id ")  WSTR_GUID_FMT _T("\n"), GUID_PRINTF_ARG(snapshotProp.m_SnapshotSetId));
													_tprintf (_T(" Provider Id ")  WSTR_GUID_FMT _T("\n"), GUID_PRINTF_ARG(snapshotProp.m_ProviderId));                         //
													_tprintf (_T(" OriginalVolumeName : %ls\n"), snapshotProp.m_pwszOriginalVolumeName);

													if (snapshotProp.m_pwszExposedName != NULL)
													{
														_tprintf (_T(" ExposedName : %ls\n"), snapshotProp.m_pwszExposedName);
													}
													if (snapshotProp.m_pwszExposedPath != NULL)
													{
														_tprintf (_T(" ExposedPath : %ls\n"), snapshotProp.m_pwszExposedPath);
													}
													if (snapshotProp.m_pwszSnapshotDeviceObject != NULL)
													{
														_tprintf (_T(" DeviceObject : %ls\n"), snapshotProp.m_pwszSnapshotDeviceObject);
													}

													/**/
													SYSTEMTIME stUTC, stLocal;
													FILETIME ftCreate;
													// Convert the last-write time to local time.
													ftCreate.dwHighDateTime  =  HILONG(snapshotProp.m_tsCreationTimestamp);
													ftCreate.dwLowDateTime   =  LOLONG(snapshotProp.m_tsCreationTimestamp);

													FileTimeToSystemTime(&ftCreate, &stUTC);
													SystemTimeToTzSpecificLocalTime(NULL, &stUTC, &stLocal);

													_tprintf (TEXT("Created : %02d/%02d/%d  %02d:%02d \n"), stLocal.wMonth, stLocal.wDay, stLocal.wYear, stLocal.wHour, stLocal.wMinute );
													_tprintf (_T("\n"));

													// Cleanup properties
													VssFreeSnapshotProperties(&snapshotProp);
												}
											}
											pDoShadowCopy->Release();
										}
									}
									pPrepare->Release();
								}
							}
						}

					}
				}
				pAsync->Release();

				//
			}
			//  free system resources allocated when IVssBackupComponents::GatherWriterMetadata was called.
			pBackup->FreeWriterMetadata();
		}
		pBackup->Release();
	}
	return returnValue;
}


void ReleaseInterface (IUnknown* unkn)
{

	if (unkn != NULL)
	{
		unkn->Release();
	}

}


