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

#define WORD_USE_NSAppleScript 1
#define WORD_USE_SBApplication 0

#define POWERPOINT_USE_NSAppleScript 1
#define POWERPOINT_USE_SBApplication 0

//only Excel seems to work as advertised by sdef

#define EXCEL_USE_NSAppleScript 1
#define EXCEL_USE_SBApplication 0

/*
 to extract sdef
 sdef Word.app > word.sdef
 sdp --basename Word -fh Word.sdef Word.h
 Note: the sdef has error for Word and PowerPoint (go figure) manually edit the sdef to get complete file
 */

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

BOOL isBundleAppSandboxed(NSBundle *b)
{
    //DEPRECATED: multiple versions may be installed; we can not depend on bundle info like thid
    if(b)
    {
        NSString *CFBundleShortVersionString = [b objectForInfoDictionaryKey:@"CFBundleShortVersionString"];
        
        if(CFBundleShortVersionString)
        {
            if([CFBundleShortVersionString length] > 1)
            {
                NSRange r = [CFBundleShortVersionString rangeOfComposedCharacterSequencesForRange:NSMakeRange(0, 2)];
                NSInteger majorVersion = [[CFBundleShortVersionString substringWithRange:r] integerValue];
                //Office 2011=14.x.x, 2016=15.x
                return (majorVersion > 14);
            }
        }
    }

    return FALSE;
}

NSString *getHFSPath(NSURL *url)
{
    NSString *path = @"";
    if(url)
    {
        NSString *_path = (NSString *)CFURLCopyFileSystemPath((CFURLRef)url, kCFURLHFSPathStyle);
        if(_path)
        {
            path = [NSString stringWithString:_path];
        }
    }
    
    return path;
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
                            BOOL isDirectory = NO;
                            if(!(([[NSFileManager defaultManager]fileExistsAtPath:[url path] isDirectory:&isDirectory]) && isDirectory))
                            {
                                [[NSFileManager defaultManager]createDirectoryAtURL:url
                                                        withIntermediateDirectories:YES
                                                                         attributes:nil
                                                                              error:nil];
                            }
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

NSURL *getAlteredURL(NSURL *folderURL, NSString *prefix, NSString *suffix)
{
    if(folderURL)
    {
        NSArray *files = [[NSFileManager defaultManager]contentsOfDirectoryAtPath:[folderURL path] error:nil];
        if(files)
        {
            // https://developer.apple.com/library/content/documentation/Cocoa/Conceptual/Predicates/Articles/pSyntax.html
            NSArray *dst_files = [files filteredArrayUsingPredicate:[NSPredicate
                                                                     predicateWithFormat:@"self BEGINSWITH %@ \
                                                                     AND self ENDSWITH %@", prefix, suffix]];
            if([dst_files count] > 0)
            {
                return [folderURL URLByAppendingPathComponent:[dst_files objectAtIndex:0]];
            }
        }
    }
    
    //write failed, for example
    return [folderURL URLByAppendingPathComponent:[NSString stringWithFormat:@"%@%@", prefix, suffix]];
}

#pragma mark -

void doItWithWord(NSURL *src_url, NSURL *dst_url)
{
	WordApplication *app = [SBApplication applicationWithBundleIdentifier:WORD_UTI];
    
    NSString *version = app.applicationVersion;
    
    BOOL isSandboxed = NO;
    
    if([version length] > 1)
    {
        NSRange r = [version rangeOfComposedCharacterSequencesForRange:NSMakeRange(0, 2)];
        NSInteger majorVersion = [[version substringWithRange:r] integerValue];
        isSandboxed = (majorVersion > 14);
    }
    
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
				
				if(isSandboxed)
				{
					tempFolderURL = getDarwinTempFoderURL(b);
				}else
				{
					tempFolderURL = getTempFoderURL();
				}
				
				if(tempFolderURL)
				{
                    NSLog(@"Scripting Microsoft Word (%@) in temporary directory:%@", version, [tempFolderURL path]);
                    
                    //Word does not have automationSecurity or alertBeforeOverwriting
					app.startupDialog = NO;
					app.displayAlerts = NO;
                    
                    NSString *src_uuid = getUUID();
                    NSString *dst_uuid = getUUID();
                    
                    NSURL *src_url_temp = [[tempFolderURL URLByAppendingPathComponent:src_uuid] URLByAppendingPathExtension:[src_url pathExtension]];
                    NSURL *dst_url_temp = [[tempFolderURL URLByAppendingPathComponent:dst_uuid] URLByAppendingPathExtension:@"pdf"];
					
					if([[NSFileManager defaultManager]copyItemAtURL:src_url toURL:src_url_temp error:nil])
					{
                        // https://msdn.microsoft.com/en-us/library/microsoft.office.tools.word.document.saveas(v=vs.120).aspx
                        // Note: "save as" fails if the file already exists

#if WORD_USE_NSAppleScript
                        NSString *script = [NSString stringWithFormat:
                                            @"on convert_to_pdf(src_path, dst_path) \n\
                                            set src_path to POSIX file src_path as string \n\
                                            set dst_path to POSIX file dst_path as string \n\
                                            tell application id \"%@\" \n\
                                            set d to open file name src_path \
                                            with read only and Revert without add to recent files and confirm conversions \n\
                                            save as d file name dst_path file format format PDF \
                                            with embed truetype fonts, save native picture format and save forms data without add to recent files \n\
                                            close d saving no saving in src_path \n\
                                            end tell  \n\
                                            end convert_to_pdf", WORD_UTI];
                        
                        NSAppleScript *scriptObject = [[NSAppleScript alloc]initWithSource:script];
                        NSAppleEventDescriptor *parameters = [NSAppleEventDescriptor listDescriptor];
                        NSAppleEventDescriptor *src_path = [NSAppleEventDescriptor descriptorWithString:[src_url_temp path]];
                        [parameters insertDescriptor:src_path atIndex:1];
                        NSAppleEventDescriptor *dst_path = [NSAppleEventDescriptor descriptorWithString:[dst_url_temp path]];
                        [parameters insertDescriptor:dst_path atIndex:2];
                        
                        ProcessSerialNumber psn = {0, kCurrentProcess};
                        NSAppleEventDescriptor *target =
                        [NSAppleEventDescriptor descriptorWithDescriptorType:typeProcessSerialNumber
                                                                       bytes:&psn
                                                                      length:sizeof(ProcessSerialNumber)];
                        
                        NSAppleEventDescriptor *handler = [NSAppleEventDescriptor descriptorWithString:@"convert_to_pdf"];
                        
                        NSAppleEventDescriptor *event =
                        [NSAppleEventDescriptor appleEventWithEventClass:kASAppleScriptSuite
                                                                 eventID:kASSubroutineEvent
                                                        targetDescriptor:target
                                                                returnID:kAutoGenerateReturnID
                                                           transactionID:kAnyTransactionID];
                        
                        [event setParamDescriptor:handler forKeyword:keyASSubroutineName];
                        [event setParamDescriptor:parameters forKeyword:keyDirectObject];
                        
                        if([scriptObject compileAndReturnError:nil])
                        {
                            NSLog(@"Opening path:%@", [src_url_temp path]);
                            NSLog(@"Saving path:%@", [dst_url_temp path]);
                            
                            [scriptObject executeAppleEvent:event error:nil];
                        }
                        [scriptObject release];
#endif
                        
#if WORD_USE_SBApplication
                        //doesn't work!
                        /*
                        NSString *src_path_temp = isSandboxed ? [src_url_temp path] : getHFSPath(src_url_temp);
                        NSString *dst_path_temp = isSandboxed ? [dst_url_temp path] : getHFSPath(dst_url_temp);
                        
                        NSLog(@"Opening path:%@", src_path_temp);
                        
                        [app openFileName:src_path_temp
                       confirmConversions:NO
                                 readOnly:YES
                         addToRecentFiles:NO
                         passwordDocument:@""
                         passwordTemplate:@""
                                   Revert:YES
                            writePassword:@""
                    writePasswordTemplate:@""
                            fileConverter:WordE162OpenFormatOffice];
                        
                        SBElementArray *documents = [app documents];
                         */
#endif
                        //does Word appends the template name to file?
                        dst_url_temp = getAlteredURL(tempFolderURL, dst_uuid, @".pdf");
                        [[NSFileManager defaultManager]moveItemAtURL:dst_url_temp toURL:dst_url error:nil];
                        [[NSFileManager defaultManager]removeItemAtURL:src_url_temp error:nil];
					}
				}
			}
		}
	}
}

void doItWithExcel(NSURL *src_url, NSURL *dst_url)
{
	ExcelApplication *app = [SBApplication applicationWithBundleIdentifier:EXCEL_UTI];

    NSString *version = [[NSNumber numberWithInteger:app.calculationVersion]stringValue];

    BOOL isSandboxed = NO;
    
    if([version length] > 1)
    {
        NSRange r = [version rangeOfComposedCharacterSequencesForRange:NSMakeRange(0, 2)];
        NSInteger majorVersion = [[version substringWithRange:r] integerValue];
        isSandboxed = (majorVersion > 14);
    }

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
				
				if(isSandboxed)
				{
					tempFolderURL = getDarwinTempFoderURL(b);
				}else
				{
					tempFolderURL = getTempFoderURL();
				}
				
				if(tempFolderURL)
				{
                    NSLog(@"Scripting Microsoft Excel (%@) in temporary directory:%@", version, [tempFolderURL path]);

					app.automationSecurity = ExcelMAScMsoAutomationSecurityLow;
					app.startupDialog = NO;
					app.alertBeforeOverwriting = NO;
					app.displayAlerts = NO;
					
                    NSString *src_uuid = getUUID();
                    NSString *dst_uuid = getUUID();
                    
					NSURL *src_url_temp = [[tempFolderURL URLByAppendingPathComponent:src_uuid] URLByAppendingPathExtension:[src_url pathExtension]];
					NSURL *dst_url_temp = [[tempFolderURL URLByAppendingPathComponent:dst_uuid] URLByAppendingPathExtension:@"pdf"];
				
					if([[NSFileManager defaultManager]copyItemAtURL:src_url toURL:src_url_temp error:nil])
					{
#if EXCEL_USE_NSAppleScript
                            NSString *script = [NSString stringWithFormat:
                                                @"on convert_to_pdf(src_path, dst_path) \n\
                                                set src_path to POSIX file src_path as string \n\
                                                set dst_path to POSIX file dst_path as string \n\
                                                tell application id \"%@\" \n\
                                                set w to open workbook workbook file name src_path \
                                                with read only and ignore read only recommended without notify and add to mru \n\
                                                try \n\
                                                save workbook as w filename dst_path file format PDF file format \
                                                conflict resolution other session changes with overwrite without read only recommended, create backup and add to most recently used list \n\
                                                end try \n\
                                                close w saving no \n\
                                                end tell \n\
                                                end convert_to_pdf", EXCEL_UTI];
                            NSAppleScript *scriptObject = [[NSAppleScript alloc]initWithSource:script];
                            NSAppleEventDescriptor *parameters = [NSAppleEventDescriptor listDescriptor];
                            NSAppleEventDescriptor *src_path = [NSAppleEventDescriptor descriptorWithString:[src_url_temp path]];
                            [parameters insertDescriptor:src_path atIndex:1];
                            NSAppleEventDescriptor *dst_path = [NSAppleEventDescriptor descriptorWithString:[dst_url_temp path]];
                            [parameters insertDescriptor:dst_path atIndex:2];
                            
                            ProcessSerialNumber psn = {0, kCurrentProcess};
                            NSAppleEventDescriptor *target =
                            [NSAppleEventDescriptor descriptorWithDescriptorType:typeProcessSerialNumber
                                                                           bytes:&psn
                                                                          length:sizeof(ProcessSerialNumber)];
                            
                            NSAppleEventDescriptor *handler = [NSAppleEventDescriptor descriptorWithString:@"convert_to_pdf"];
                            
                            NSAppleEventDescriptor *event =
                            [NSAppleEventDescriptor appleEventWithEventClass:kASAppleScriptSuite
                                                                     eventID:kASSubroutineEvent
                                                            targetDescriptor:target
                                                                    returnID:kAutoGenerateReturnID
                                                               transactionID:kAnyTransactionID];
                            
                            [event setParamDescriptor:handler forKeyword:keyASSubroutineName];
                            [event setParamDescriptor:parameters forKeyword:keyDirectObject];
                            
                            if([scriptObject compileAndReturnError:nil])
                            {
                                NSLog(@"Opening path:%@", [src_url_temp path]);
                                NSLog(@"Saving path:%@", [dst_url_temp path]);
                                
                                [scriptObject executeAppleEvent:event error:nil];
                            }
                            [scriptObject release];
                        
#endif

#if EXCEL_USE_SBApplication
                        //Excel 2011 wants HFS; Excel 2016 accepts POSIX too
                        NSString *src_path_temp = isSandboxed ? [src_url_temp path] : getHFSPath(src_url_temp);
                        NSString *dst_path_temp = isSandboxed ? [dst_url_temp path] : getHFSPath(dst_url_temp);
                        
                        NSLog(@"Opening path:%@", src_path_temp);
                        
                        ExcelWorkbook *doc = [app openWorkbookWorkbookFileName:src_path_temp
                                                                   updateLinks:ExcelE294DoNotUpdateLinks
                                                                      readOnly:YES
                                                                        format:ExcelE295CommasDelimiter
                                                                      password:@""
                                                         writeReservedPassword:@""
                                                     ignoreReadOnlyRecommended:NO
                                                                        origin:ExcelE211Macintosh
                                                                     delimiter:@""
                                                                      editable:NO
                                                                        notify:NO
                                                                     converter:0
                                                                      addToMru:NO];
						if(doc)
						{
                            NSLog(@"Saving path:%@", dst_path_temp);
                            
                            [doc saveWorkbookAsFilename:dst_path_temp
                                             fileFormat:ExcelE188PDFFileFormat
                                               password:@""
                               writeReservationPassword:@""
                                    readOnlyRecommended:NO
                                           createBackup:NO
                                             accessMode:ExcelE221NoChange
                                     conflictResolution:ExcelE222OtherSessionChanges
                              addToMostRecentlyUsedList:NO
                                              overwrite:YES];
                            
                            [doc closeSaving:ExcelSavoNo savingIn:src_url_temp];

						}//doc
#endif
                        //Excel appends the template name to file
                        dst_url_temp = getAlteredURL(tempFolderURL, dst_uuid, @".pdf");
                        [[NSFileManager defaultManager]moveItemAtURL:dst_url_temp toURL:dst_url error:nil];
                        [[NSFileManager defaultManager]removeItemAtURL:src_url_temp error:nil];
					}
				}
			}
		}
	}
}

void doItWithPowerPoint(NSURL *src_url, NSURL *dst_url)
{
	PowerPointApplication *app = [SBApplication applicationWithBundleIdentifier:POWERPOINT_UTI];
	
    NSString *version = app.Version;
    
    BOOL isSandboxed = NO;
    
    if([version length] > 1)
    {
        NSRange r = [version rangeOfComposedCharacterSequencesForRange:NSMakeRange(0, 2)];
        NSInteger majorVersion = [[version substringWithRange:r] integerValue];
        isSandboxed = (majorVersion > 14);
    }
    
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
				
				if(isSandboxed)
				{
					tempFolderURL = getDarwinTempFoderURL(b);
				}else
				{
					tempFolderURL = getTempFoderURL();
				}
				
				if(tempFolderURL)
				{
                    NSLog(@"Scripting Microsoft PowerPoint (%@) in temporary directory:%@", version, [tempFolderURL path]);
                    
                    //PowerPoint doesn't have automationSecurity, alertBeforeOverwriting or displayAlerts
					app.startUpDialog = NO;
					
                    NSString *src_uuid = getUUID();
                    NSString *dst_uuid = getUUID();
                    
                    NSURL *src_url_temp = [[tempFolderURL URLByAppendingPathComponent:src_uuid] URLByAppendingPathExtension:[src_url pathExtension]];
                    NSURL *dst_url_temp = [[tempFolderURL URLByAppendingPathComponent:dst_uuid] URLByAppendingPathExtension:@"pdf"];

					if([[NSFileManager defaultManager]copyItemAtURL:src_url toURL:src_url_temp error:nil])
					{
#if POWERPOINT_USE_NSAppleScript
                        NSString *script = [NSString stringWithFormat:
                                            @"on convert_to_pdf(src_path, dst_path) \n\
                                            set src_path to POSIX file src_path as string \n\
                                            set dst_path to POSIX file dst_path as string \n\
                                            tell application id \"%@\" \n\
                                            open src_path \n\
                                            set pp to presentations whose full name is src_path \n\
                                            if (count pp) is 1 then \n\
                                            set p to item 1 of pp \n\
                                            save p in dst_path as save as PDF \n\
                                            close p saving no saving in src_path \n\
                                            end if \n\
                                            end tell \n\
                                            end convert_to_pdf", POWERPOINT_UTI];
                        NSAppleScript *scriptObject = [[NSAppleScript alloc]initWithSource:script];
                        NSAppleEventDescriptor *parameters = [NSAppleEventDescriptor listDescriptor];
                        NSAppleEventDescriptor *src_path = [NSAppleEventDescriptor descriptorWithString:[src_url_temp path]];
                        [parameters insertDescriptor:src_path atIndex:1];
                        NSAppleEventDescriptor *dst_path = [NSAppleEventDescriptor descriptorWithString:[dst_url_temp path]];
                        [parameters insertDescriptor:dst_path atIndex:2];
                        
                        ProcessSerialNumber psn = {0, kCurrentProcess};
                        NSAppleEventDescriptor *target =
                        [NSAppleEventDescriptor descriptorWithDescriptorType:typeProcessSerialNumber
                                                                       bytes:&psn
                                                                      length:sizeof(ProcessSerialNumber)];
                        
                        NSAppleEventDescriptor *handler = [NSAppleEventDescriptor descriptorWithString:@"convert_to_pdf"];
                        
                        NSAppleEventDescriptor *event =
                        [NSAppleEventDescriptor appleEventWithEventClass:kASAppleScriptSuite
                                                                 eventID:kASSubroutineEvent
                                                        targetDescriptor:target
                                                                returnID:kAutoGenerateReturnID
                                                           transactionID:kAnyTransactionID];
                        
                        [event setParamDescriptor:handler forKeyword:keyASSubroutineName];
                        [event setParamDescriptor:parameters forKeyword:keyDirectObject];
                        
                        if([scriptObject compileAndReturnError:nil])
                        {
                            NSLog(@"Opening path:%@", [src_url_temp path]);
                            NSLog(@"Saving path:%@", [dst_url_temp path]);
                            
                            [scriptObject executeAppleEvent:event error:nil];
                        }
                        [scriptObject release];
#endif
                        dst_url_temp = getAlteredURL(tempFolderURL, dst_uuid, @".pdf");
                        [[NSFileManager defaultManager]moveItemAtURL:dst_url_temp toURL:dst_url error:nil];
                        [[NSFileManager defaultManager]removeItemAtURL:src_url_temp error:nil];
					}
				}
			}
		}
	}
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
        NSString *uti = nil;
        
        if ([src_url getResourceValue:&uti forKey:NSURLTypeIdentifierKey error:nil])
        {
            if ([[NSWorkspace sharedWorkspace] type:uti conformsToType:@"org.openxmlformats.spreadsheetml.sheet"])
            {
                doItWithExcel(src_url, dst_url);
            }else
                if ([[NSWorkspace sharedWorkspace] type:uti conformsToType:@"org.openxmlformats.presentationml.presentation"])
                {
                    doItWithPowerPoint(src_url, dst_url);
                }else
                    if ([[NSWorkspace sharedWorkspace] type:uti conformsToType:@"org.openxmlformats.wordprocessingml.document"])
                    {
                        doItWithWord(src_url, dst_url);
                    }
        }
	}
	
	[src_url release];
	[dst_url release];
}

