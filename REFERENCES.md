# References

| Artifact Type | URL |
| --- | --- |
| ESE DB Format | <https://github.com/digitalsleuth/libesedb/blob/main/documentation/Extensible%20Storage%20Engine%20(ESE)%20Database%20File%20(EDB)%20format.asciidoc> |
| Media Formats | <https://gix.github.io/media-types/> |
| Property Keys | <https://learn.microsoft.com/en-us/windows/win32/properties/props><br><br><https://learn.microsoft.com/en-us/windows/win32/search/-search-3x-wds-propertymappings><br><br><https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wsp/67328dcc-4e12-4e1e-be80-d91684df2f98><br><br><https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wsp/2dbe759c-c955-4770-a545-e46d7f6332ed><br><br><https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/bf7aeae8-c47a-4939-9f45-700158dac3bc><br><br><https://cdn.callback.com/shellboost/doc/Reference/ShellBoost-Core-Assembly/html/f2805716-6023-897c-2234-d79a4feb5c5f.htm> |
| Serialized Property Store (SPS) | <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-propstore/3453fb82-0e4f-4c2c-bc04-64b4bd2c51ec><br><br><https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/f122b9d7-e5cf-4484-8466-83f6fd94b3cc><br><br><https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/5a2b34c4-d109-438e-9ec8-84816d8de40d> (0x06 Currency)<br><br><https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/04727953-c174-4d01-80ae-b38b15d7068a> (CodePage Property) |
| Variant Types | <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f><br><br><https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.package.variant.varianttype?view=visualstudiosdk-2019><br><br><https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum> |
| 7-bit compression | <https://www.3gpp.org/ftp/Specs/archive/23_series/23.038/><br><br><https://doubleblak.com/blogPost.php?k=7bitpdu> |
| ESE multi-value columns (libesedb API) | <https://github.com/libyal/libesedb/> |
| Legacy Windows.edb text encoding (AutoSummary obfuscation and compression, Music.Artist / Music.Genre) | <https://github.com/libyal/documentation/blob/master/Forensic%20analysis%20of%20the%20Windows%20Search%20database.pdf><br><br><https://eprints.whiterose.ac.uk/75046/1/Forensic_Data_Recovery_From_The_Windows_Search_Database_preprint_DIIN328.pdf><br><br><https://www.dataforensics.org/?p=1144><br><br><https://github.com/libyal/libesedb/blob/main/esedbtools/windows_search.c><br><br><https://github.com/Sootmark/search/blob/main/src/encoded.rs><br><br><https://github.com/Sootmark/search/blob/main/src/property.rs> |
| mapi16 URLs (Hangul-encoded IDs, store hash) | <https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/about-mapi-urls-for-notification-based-indexing><br><br><https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/algorithm-to-encode-entry-ids-and-attachment-ids><br><br><https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/algorithm-to-calculate-the-store-hash-number><br><br><https://github.com/microsoft/mfcmapi> (`core\smartview\encodeEntryID.cpp`, `decodeEntryID.cpp`, `ComputeStoreHash`) |
| Outlook entry ID layout (message and attachment IDs) | <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/141923d5-15ab-4ef1-a524-6dce75aae546> (section 2.4.3.2) |
| System.Message.ConversationIndex | <https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/daa9120f-f325-4afb-a738-28f91049ab3c> (section 2.2.1.3) |
| MAPI property names (PidTag / PidLid / PidName) | <https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/f6ab1613-aefe-447d-a49c-18217230b148> |
| GKPIDDSI document summary property names | <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/d93502fa-5b8f-4f47-a3fe-5574046f4b8d> |
| VT_ERROR (HRESULT) | <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/a9046ed2-bfb2-4d56-a719-2824afce59ac><br><br><https://learn.microsoft.com/en-us/windows/win32/search/-search-prth-error-constants> |
| System.Calendar.ResponseStatus | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-calendar-responsestatus> |
| System.Calendar.ShowTimeAs | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-calendar-showtimeas> |
| System.Capacity | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-capacity> |
| System.Communication.TaskStatus | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-communication-taskstatus> |
| System.Document.TotalEditingTime | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-document-totaleditingtime> |
| System.FileAttributes | <https://learn.microsoft.com/en-us/windows/win32/fileio/file-attribute-constants> |
| System.FilePlaceholderStatus | <https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/ne-shobjidl_core-placeholder_states> |
| System.FlagColor | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-flagcolor> |
| System.FlagStatus | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-flagstatus> |
| System.Image.Compression | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-image-compression> |
| System.Importance | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-importance> |
| System.Link.TargetSFGAOFlags<br>System.SFGAOFlags | <https://learn.microsoft.com/en-us/windows/win32/shell/sfgao> |
| System.Media.Duration | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-media-duration> |
| System.Message.Flags | <https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagmessageflags-canonical-property><br><br><https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OXCMSG/%5bMS-OXCMSG%5d.pdf> |
| System.Sensitivity | <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-sensitivity> |
| System.SyncTransferStatusFlags | <https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/ne-shobjidl_core-sync_transfer_status> |
