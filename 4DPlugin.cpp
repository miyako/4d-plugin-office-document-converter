/* --------------------------------------------------------------------------------
 #
 #	4DPlugin.cpp
 #	source generated by 4D Plugin Wizard
 #	Project : Office
 #	author : miyako
 #	2017/10/14
 #
 # --------------------------------------------------------------------------------*/


#include "4DPluginAPI.h"
#include "4DPlugin.h"

void PluginMain(PA_long32 selector, PA_PluginParameters params)
{
	try
	{
		PA_long32 pProcNum = selector;
		sLONG_PTR *pResult = (sLONG_PTR *)params->fResult;
		PackagePtr pParams = (PackagePtr)params->fParameters;

		CommandDispatcher(pProcNum, pResult, pParams); 
	}
	catch(...)
	{

	}
}

void CommandDispatcher (PA_long32 pProcNum, sLONG_PTR *pResult, PackagePtr pParams)
{
	switch(pProcNum)
	{
// --- Office

		case 1 :
			OFFICE_DOCUMENT_TO_PDF(pResult, pParams);
			break;

	}
}

#pragma mark -

// ------------------------------------ Office ------------------------------------

NSBundle *getBundle(NSString *uti, NSString *name)
{
	//NSBundle fails to find the application by its uti; need to use NSWorkspace instead
	NSString *path = nil;
	NSURL *url = [[NSWorkspace sharedWorkspace]URLForApplicationWithBundleIdentifier:uti];
	
	if(url)
	{
		path = [url path];
	}else
	{
		//maybe the uti changed? try again with name
		path = [[NSWorkspace sharedWorkspace]fullPathForApplication:name];
		if(path)
		{
			url = [NSURL fileURLWithPath:path];
		}
	}
	
	if(url)
	{
		return [NSBundle bundleWithURL:url];
	}
	
	return nil;
}

BOOL isBundleAppSandboxed(NSBundle *b)
{
	if(b)
	{
		NSString *MicrosoftBuildNumber = [b objectForInfoDictionaryKey:@"MicrosoftBuildNumber"];
		
		if(MicrosoftBuildNumber)
		{
			if([MicrosoftBuildNumber length] > 1)
			{
				NSRange r = [MicrosoftBuildNumber rangeOfComposedCharacterSequencesForRange:NSMakeRange(0, 2)];
				NSInteger majorVersion = [[MicrosoftBuildNumber substringWithRange:r] integerValue];
				
				return (majorVersion > 14);
			}
		}
	}
	return FALSE;
}

NSURL *getAppContainerURL(NSBundle *b)
{
	NSURL *url = nil;
	
	if(b)
	{
		NSArray *URLs = [[NSFileManager defaultManager]
										 URLsForDirectory:NSLibraryDirectory
										 inDomains:NSUserDomainMask];
		if(URLs && [URLs count])
		{
			NSURL *libraryURL = [URLs objectAtIndex:0];
			if(libraryURL)
			{
				NSURL *containerURL = [libraryURL URLByAppendingPathComponent:@"Containers"];
				url = [containerURL URLByAppendingPathComponent:[b bundleIdentifier] isDirectory:YES];
			}
		}
	}
	
	return url;
}

NSURL *getDarwinTempFoderURL(NSBundle *b)
{
	NSURL *url = getAppContainerURL(b);
	
	if(url)
	{
		NSURL *infoURL = [url URLByAppendingPathComponent:@"Container.plist"];
		
		if(infoURL)
		{
			NSLog(@"container for %@:%@\n",[b bundleIdentifier], [infoURL path]);
			NSDictionary *plist = [[NSDictionary alloc]initWithContentsOfURL:infoURL];
			if(plist)
			{
				NSDictionary *SandboxProfileDataValidationInfo = [plist objectForKey:@"SandboxProfileDataValidationInfo"];
				if(SandboxProfileDataValidationInfo)
				{
					NSDictionary *SandboxProfileDataValidationParametersKey = [SandboxProfileDataValidationInfo objectForKey:@"SandboxProfileDataValidationParametersKey"];
					if(SandboxProfileDataValidationParametersKey)
					{
						NSString *application_darwin_temp_dir = [SandboxProfileDataValidationParametersKey objectForKey:@"application_darwin_temp_dir"];
						if(application_darwin_temp_dir)
						{
							url = [NSURL fileURLWithPath:application_darwin_temp_dir];
						}
					}
				}
				[plist release];
			}
		}
	}
	
	return url;
}

NSURL *getTempFoderURL()
{
	NSURL *url = nil;
	
	NSArray *URLs = [[NSFileManager defaultManager]
									 URLsForDirectory:NSDesktopDirectory
									 inDomains:NSUserDomainMask];
	
	if(URLs && [URLs count])
	{
		NSURL *desktopURL = [URLs objectAtIndex:0];
		url = [[NSFileManager defaultManager]
					 URLForDirectory:NSItemReplacementDirectory
					 inDomain:NSUserDomainMask
					 appropriateForURL:desktopURL
					 create:YES
					 error:nil];
	}
	
	return url;
}

NSString *getUUID()
{
	return [[[NSUUID UUID]UUIDString]stringByReplacingOccurrencesOfString:@"-" withString:@""];
}

#pragma mark -

BOOL doItWithWord(NSURL *src_url, NSURL *dst_url)
{
	BOOL success = NO;
	
	WordApplication *app = [SBApplication applicationWithBundleIdentifier:WORD_UTI];
	NSURL *appUrl = [[NSWorkspace sharedWorkspace]URLForApplicationWithBundleIdentifier:WORD_UTI];
	
	Boolean accepts = false;
	if(noErr == LSCanURLAcceptURL((CFURLRef)src_url, (CFURLRef)appUrl, kLSRolesAll, kLSAcceptDefault, &accepts))
	{
		if(accepts)
		{
			NSBundle *b = getBundle(WORD_UTI, WORD_NAME);
			
			if(b)
			{
				NSURL *tempFolderURL = nil;
				BOOL isSandboxed = isBundleAppSandboxed(b);
				
				if(isSandboxed)
				{
					tempFolderURL = getDarwinTempFoderURL(b);
				}else
				{
					tempFolderURL = getTempFoderURL();
				}
				
				if(tempFolderURL)
				{
					app.automationSecurity = WordAutomationSecurityAlertsNone;
					app.startupDialog = NO;
					app.displayAlerts = NO;
					
					NSURL *src_url_temp = [[tempFolderURL URLByAppendingPathComponent:getUUID()] URLByAppendingPathExtension:[src_url pathExtension]];
					NSURL *dst_url_temp = [[tempFolderURL URLByAppendingPathComponent:getUUID()] URLByAppendingPathExtension:@".pdf"];
					
					if([[NSFileManager defaultManager]copyItemAtURL:src_url toURL:src_url_temp error:nil])
					{
						WordDocument *doc = [app openFileName:[src_url_temp path]
															 confirmConversions:NO
																				 readOnly:YES
																 addToRecentFiles:NO
																 passwordDocument:@""
																 passwordTemplate:@""
																					 Revert:NO
																		writePassword:@""
														writePasswordTemplate:@""
																		fileConverter:WordE162OpenFormatAuto];
						if(doc)
						{
							[doc saveAsFileName:[dst_url_temp path]
											 fileFormat:WordFileFormatFormatPDF
										 lockComments:NO
												 password:@""
								 addToRecentFiles:NO
										writePassword:@""
							readOnlyRecommended:NO
							 embedTruetypeFonts:YES
					saveNativePictureFormat:YES
										saveFormsData:YES
										 textEncoding:WordMtEnEncodingUTF8
								 insertLineBreaks:YES
							 allowSubstitutions:YES
									 lineEndingType:WordLineEndingTypeLineEndingCrLf
						HTMLDisplayOnlyOutput:NO
						maintainCompatibility:NO];
							
							[doc closeSaving:WordSaveOptionNo savingIn:dst_url_temp];
							[[NSFileManager defaultManager]copyItemAtURL:dst_url_temp toURL:dst_url error:nil];
						}
						[[NSFileManager defaultManager]removeItemAtURL:src_url_temp error:nil];
						[[NSFileManager defaultManager]removeItemAtURL:dst_url_temp error:nil];
					}
				}
			}
		}
	}
	
	return success;
}
BOOL doItWithExcel(NSURL *src_url, NSURL *dst_url)
{
	BOOL success = NO;

	ExcelApplication *app = [SBApplication applicationWithBundleIdentifier:EXCEL_UTI];
	NSURL *appUrl = [[NSWorkspace sharedWorkspace]URLForApplicationWithBundleIdentifier:EXCEL_UTI];
	
	Boolean accepts = false;
	if(noErr == LSCanURLAcceptURL((CFURLRef)src_url, (CFURLRef)appUrl, kLSRolesAll, kLSAcceptDefault, &accepts))
	{
		if(accepts)
		{
			NSBundle *b = getBundle(EXCEL_UTI, EXCEL_NAME);
			
			if(b)
			{
				NSURL *tempFolderURL = nil;
				BOOL isSandboxed = isBundleAppSandboxed(b);
				
				if(isSandboxed)
				{
					tempFolderURL = getDarwinTempFoderURL(b);
				}else
				{
					tempFolderURL = getTempFoderURL();
				}
				
				if(tempFolderURL)
				{
					app.automationSecurity = ExcelMAScMsoAutomationSecurityLow;
					app.startupDialog = NO;
					app.alertBeforeOverwriting = NO;
					app.displayAlerts = NO;
					
					NSURL *src_url_temp = [[tempFolderURL URLByAppendingPathComponent:getUUID()] URLByAppendingPathExtension:[src_url pathExtension]];
					NSURL *dst_url_temp = [[tempFolderURL URLByAppendingPathComponent:getUUID()] URLByAppendingPathExtension:@".pdf"];
				
					if([[NSFileManager defaultManager]copyItemAtURL:src_url toURL:src_url_temp error:nil])
					{
						ExcelWorkbook *doc = [app openWorkbookWorkbookFileName:[src_url_temp path]
																											 updateLinks:ExcelUpdateLinksOptionDoNotUpdateLinks
																													readOnly:YES
																														format:ExcelFormatCommasDelimiter
																													password:@""
																						 writeReservedPassword:@""
																				 ignoreReadOnlyRecommended:NO
																														origin:ExcelOriginMacintosh
																												 delimiter:@""
																													editable:NO
																														notify:NO
																												 converter:0
																													addToMru:NO];
						if(doc)
						{
							[doc saveWorkbookAsFilename:[dst_url_temp path]
															 fileFormat:ExcelFileFormatPDFFileFormat
																 password:@""
								 writeReservationPassword:@""
											readOnlyRecommended:NO
														 createBackup:NO
															 accessMode:ExcelAccessModeNoChange
											 conflictResolution:ExcelConflictResolutionOtherSessionChanges
								addToMostRecentlyUsedList:NO
																overwrite:YES];
							
							[doc closeSaving:ExcelSaveOptionNo savingIn:dst_url_temp];
							[[NSFileManager defaultManager]copyItemAtURL:dst_url_temp toURL:dst_url error:nil];
						}
						[[NSFileManager defaultManager]removeItemAtURL:src_url_temp error:nil];
						[[NSFileManager defaultManager]removeItemAtURL:dst_url_temp error:nil];
					}
				}
			}
		}
	}
	
	return success;
}

BOOL doItWithPowerPoint(NSURL *src_url, NSURL *dst_url)
{
	BOOL success = NO;
	
	PowerPointApplication *app = [SBApplication applicationWithBundleIdentifier:POWERPOINT_UTI];
	NSURL *appUrl = [[NSWorkspace sharedWorkspace]URLForApplicationWithBundleIdentifier:POWERPOINT_UTI];
	
	Boolean accepts = false;
	if(noErr == LSCanURLAcceptURL((CFURLRef)src_url, (CFURLRef)appUrl, kLSRolesAll, kLSAcceptDefault, &accepts))
	{
		if(accepts)
		{
			NSBundle *b = getBundle(POWERPOINT_UTI, POWERPOINT_NAME);
			
			if(b)
			{
				NSURL *tempFolderURL = nil;
				BOOL isSandboxed = isBundleAppSandboxed(b);
				
				if(isSandboxed)
				{
					tempFolderURL = getDarwinTempFoderURL(b);
				}else
				{
					tempFolderURL = getTempFoderURL();
				}
				
				if(tempFolderURL)
				{
					app.automationSecurity = PowerPointAutomationSecurityMsoAutomationSecurityLow;
					app.startUpDialog = NO;
					
					NSURL *src_url_temp = [[tempFolderURL URLByAppendingPathComponent:getUUID()] URLByAppendingPathExtension:[src_url pathExtension]];
					NSURL *dst_url_temp = [[tempFolderURL URLByAppendingPathComponent:getUUID()] URLByAppendingPathExtension:@".pdf"];
					
					if([[NSFileManager defaultManager]copyItemAtURL:src_url toURL:src_url_temp error:nil])
					{
						PowerPointPresentation *doc = [app open:[src_url_temp path]
																					 password:@""
																			writePassword:@""];
						if(doc)
						{
							[doc saveIn:[dst_url_temp path]
											 as:PowerPointFileFormatSaveAsPDF];
							
							[doc closeSaving:PowerPointSaveOptionNo savingIn:dst_url_temp];
							[[NSFileManager defaultManager]copyItemAtURL:dst_url_temp toURL:dst_url error:nil];
						}
						[[NSFileManager defaultManager]removeItemAtURL:src_url_temp error:nil];
						[[NSFileManager defaultManager]removeItemAtURL:dst_url_temp error:nil];
					}
				}
			}
		}
	}
	
	return success;
}

#pragma mark -

void OFFICE_DOCUMENT_TO_PDF(sLONG_PTR *pResult, PackagePtr pParams)
{
	C_TEXT Param1;
	C_TEXT Param2;

	Param1.fromParamAtIndex(pParams, 1);
	Param2.fromParamAtIndex(pParams, 2);

	NSURL *src_url = Param1.copyUrl();
	NSURL *dst_url = Param2.copyUrl();
	
	@autoreleasepool
	{
		if(!doItWithExcel(src_url, dst_url))
		{
			if(!doItWithWord(src_url, dst_url))
			{
				doItWithPowerPoint(src_url, dst_url);
			}
		}
	}
	
	[src_url release];
	[dst_url release];
}

