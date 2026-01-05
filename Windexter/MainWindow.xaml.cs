using System.Collections;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using OfficeOpenXml;
using SQLitePCL;

namespace Windexter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// TODO:
    /// - Map PropMap.db Id's to .db Metadata and .edb PropertyStore Columns
    /// - HRESULT for VT_ERROR https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/a9046ed2-bfb2-4d56-a719-2824afce59ac
    ///   and https://learn.microsoft.com/en-us/windows/win32/search/-search-prth-error-constants

    public partial class MainWindow : Window
    {

        private static readonly DispatcherTimer? elapsedTimer = new();
        private static readonly Stopwatch? stopWatch = new();
        private static readonly string displayName = "Windexter";
        private static readonly string githubBinaryRepo = "https://github.com/digitalsleuth/windexter";
        private static readonly Version? appVersion = new(Assembly.GetExecutingAssembly().GetName().Version!.ToString(3));
        private static string dbFile = "";
        private static string dbType = "";
        private static string outputFile = "";
        private static List<List<object>> rows = [];
        private static readonly List<List<string>> paths = [];
        private static Dictionary<int, string> resolvedPaths = [];
        private static readonly List<List<object>> GatherResults = [];
        private static bool GatherAvailable = false;
        private static bool PropMapAvailable = false;
        private static List<List<object>> IndexResults = [];
        //private static List<List<object>> IndexProperties = [];
        private static readonly List<List<object>> URLResults = [];
        private static readonly List<List<object>> GPSResults = [];
        private static readonly List<List<object>> SummaryResults = [];
        private static readonly List<List<object>> CompInfoResults = [];
        private static readonly List<List<object>> ActivityResults = [];
        private static readonly List<List<object>> EseResults = [];
        private static List<List<object>> Timeline = [];
        private static readonly List<List<object>> PropertyMapResults = [];
        private static readonly List<string> tables = [];
        private static List<List<object>> properties = [];
        private static List<List<object>> propertyStore = [];
        private static List<List<object>> propertyMetadata = [];
        private static List<List<object>> propertyMap = [];
        private static readonly Dictionary<string, string> eseProps = [];
        private static readonly List<string> lookups = [
            "System.FlagColor",
            "System.Calendar.ResponseStatus",
            "System.Calendar.ShowTimeAs",
            "System.Sensitivity",
            "System.FlagStatus",
            "System.Image.Compression",
            "System.SyncTransferStatusFlags",
            "System.Importance",
            "System.Communication.TaskStatus"
            ];

        private static readonly List<string> dateTimes = [
            "System.ActivityHistory.StartTime",
            "System.ActivityHistory.EndTime",
            "System.ActivityHistory.LocalEndTime",
            "System.ActivityHistory.LocalStartTime",
            "System.Calendar.ReminderTime",
            "System.Communication.DateItemExpires",
            "System.Contact.Anniversary",
            "System.Contact.Birthday",
            "System.DateAccessed",
            "System.DateAcquired",
            "System.DateArchived",
            "System.DateCompleted",
            "System.DateCreated",
            "System.DateImported",
            "System.DateModified",
            "System.Document.DateCreated",
            "System.Document.DatePrinted",
            "System.Document.DateSaved",
            "System.DueDate",
            "System.EndDate",
            "System.GPS.Date",
            "System.ItemDate",
            "System.Link.DateVisited",
            "System.Media.DateEncoded",
            "System.Media.DateReleased",
            "System.Message.DateSent",
            "System.Message.DateReceived",
            "System.Photo.DateTaken",
            "System.RecordedTV.DateContentExpires",
            "System.RecordedTV.OriginalBroadcastDate",
            "System.RecordedTV.RecordingTime",
            "System.Search.GatherTime",
            "System.Software.DateLastUsed",
            "System.StartDate",
            "System.VersionControl.LastChangeDate",
            "LastModified",
            ];

        private static readonly List<string> unicodeField = [
            "System.Kind",
            "System.Author",
            "System.ItemParticipants",
            "System.ItemAuthors",
            "System.Media.DlnaProfileID",
            "System.Message.FromAddress",
            "System.Message.FromName",
            "System.Message.ToAddress",
            "System.Message.ToName",
            "System.Message.CcAddress",
            "System.Message.CcName",
            "System.Message.BccName",
            "System.Message.BccAddress",
            "System.Message.AttachmentNames",
            "System.ActivityHistory.DaysActive",
            "System.ActivityHistory.HoursActive",
            "System.Activity.AppIdList",
            "System.Media.Producer",
            "System.Contact.IMAddress",
            "System.Photo.TagViewAggregate",
            "System.Keywords",
            "System.Category",
            "System.Calendar.OptionalAttendeeNames",
            "System.Calendar.OptionalAttendeeAddresses",
            "System.Calendar.RequiredAttendeeNames",
            "System.Calendar.RequiredAttendeeAddresses",
            "System.LowKeywords",
            "System.MediumKeywords",
            "System.HighKeywords",
            ];

        private static readonly List<string> booleanField = [
            "System.IsAttachment",
            "System.IsFolder",
            "System.IsEncrypted",
            "System.IsDeleted",
            "System.Message.HasAttachments",
            "System.IsFlaggedComplete",
            "System.IsFlagged",
            "System.NotUserContent",
            "System.ActivityHistory.IsHistoryAttributedToSetAnchor",
            "System.ActivityHistory.IsLocal",
            "System.Activity.HasAdaptiveContent",
            "System.HasAdaptiveContent",
            "System.Calendar.IsRecurring",
            "System.Calendar.IsOnline",
            "System.IsRead",
            "System.DRM.IsProtected",
            "System.Video.IsStereo",
            "System.Video.IsSpherical"
            ];

        private static readonly List<string> floatField = [
            "System.GPS.LatitudeDecimal",
            "System.GPS.LongitudeDecimal",
            "System.Image.HorizontalResolution",
            "System.Image.VerticalResolution",
            "System.Photo.Aperture",
            "System.Photo.DigitalZoom",
            "System.Photo.ExposureBias",
            "System.Photo.ExposureTime",
            "System.Photo.FNumber",
            "System.Photo.FocalLength",
            "System.Photo.MaxAperture",
            "System.Photo.ShutterSpeed",
            "System.Photo.SubjectDistance",
            "System.Search.LastIndexedTotalTime",
            ];

        private static readonly List<string> bytesToString = [
            "InvertedOnlyMD5",
            "System.ThumbnailCacheId",
            ];

        private static readonly List<string> sfgaoField = [
            "System.Link.TargetSFGAOFlags", 
            "System.SFGAOFlags"
            ];

        private static readonly List<string> guids = [
            "System.Activity.ActivityId", 
            "System.ActivityHistory.Id", 
            "System.VolumeId", 
            "System.Setting.HostID",
            "System.Setting.SettingID",
            "System.Setting.PageID",
            "System.Message.ConversationID",
            ];

        private static readonly List<string> uint64List = [
            "System.ActivityHistory.Importance", 
            "System.ActivityHistory.ActiveDuration"
            ];

        private static readonly List<string> durations = [
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-document-totaleditingtime
            "System.Document.TotalEditingTime", 
            "System.Media.Duration"
            ];
        

        private static readonly Dictionary<int, string> FlagColor = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-flagcolor
            {
                {1, "Purple" },
                {2, "Orange" },
                {3, "Green" },
                {4, "Yellow" },
                {5, "Blue" },
                {6, "Red" }
            };

        private static readonly Dictionary<int, string> ResponseStatus = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-calendar-responsestatus
            {
                {0, "None" },
                {1, "Organized" },
                {2, "Tentative" },
                {3, "Accepted" },
                {4, "Declined" },
                {5, "Not Responded" }
            };

        private static readonly Dictionary<int, string> ShowTimeAs = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-calendar-showtimeas
            {
                {0, "Free" },
                {1, "Tentative" },
                {2, "Busy" },
                {3, "Out of Office" }
            };

        private static readonly Dictionary<int, string> Sensitivity = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-sensitivity
            {
                {0, "Normal" },
                {1, "Personal" },
                {2, "Private" },
                {3, "Confidential" }
            };

        private static readonly Dictionary<int, string> FlagStatus = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-flagstatus
            {
                {0, "Not Flagged" },
                {1, "Completed" },
                {2, "Follow Up" }
            };

        private static readonly Dictionary<int, string> Importance = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-importance
            {
                {0, "Low" },
                {1, "Low" },
                {2, "Normal" },
                {3, "Normal" },
                {4, "Normal" },
                {5, "High" }
            };

        private static readonly Dictionary<int, string> ImageCompression = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-image-compression
            {
                {1, "Uncompressed" },
                {2, "CCITT T.3" },
                {3, "CCITT T.4" },
                {4, "CCITT T.6" },
                {5, "LZW" },
                {6, "JPEG" },
                {32773, "PACKBITS" }
            };

        private static readonly Dictionary<int, string> SyncTransferStatus = new()
        /// https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/ne-shobjidl_core-sync_transfer_status
            {
                {0, "None" },
                {0x1, "Needs Upload" },
                {0x2, "Needs Download" },
                {0x4, "Transferring" },
                {0x8, "Paused" },
                {0x10, "Has Error" },
                {0x20, "Fetching Metadata" },
                {0x40, "User Requested Refresh" },
                {0x80, "Has Warning" },
                {0x100, "Excluded" },
                {0x200, "Incomplete" },
                {0x400, "Placeholder If Empty" }
            };

        private static readonly Dictionary<int, string> TaskStatus = new()
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-communication-taskstatus
            {
                {0, "Not Started" },
                {1, "In Progress" },
                {2, "Complete" },
                {3, "Waiting" },
                {4, "Deferred" }
            };

        private static readonly Dictionary<long, string> MessageFlags = new()
        ///https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagmessageflags-canonical-property
        ///https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OXCMSG/%5bMS-OXCMSG%5d.pdf
            {
                { 0x01, "MSGFLAG_READ"},
                { 0x02, "MSGFLAG_UNMODIFIED"},
                { 0x04, "MSGFLAG_SUBMIT"},
                { 0x08, "MSGFLAG_UNSENT"},
                { 0x10, "MSGFLAG_HASATTACH"},
                { 0x20, "MSGFLAG_FROMME"},
                { 0x40, "MSGFLAG_FAI" },
                { 0x80, "MSGFLAG_RESEND"},
                { 0x100, "MSGFLAG_NOTIFYREAD" },
                { 0x200, "MSGFLAG_NOTIFYUNREAD" },
                { 0x400, "MSGFLAG_EVERREAD" },
                { 0x2000, "MSGFLAG_ORIGIN_INTERNET"},
                { 0x8000, "MSGFLAG_UNTRUSTED" },
                ///MSGFLAG_ASSOCIATED
                ///MSGFLAG_NRN_PENDING
                ///MSGFLAG_ORIGIN_MISC_EXT
                ///MSGFLAG_ORIGIN_X400
                ///MSGFLAG_ORIGIN_EXT_SEND
                ///MSGFLAG_RN_PENDING
            };

        private static readonly Dictionary<long, string> SFGAO = new() 
        /// https://learn.microsoft.com/en-us/windows/win32/shell/sfgao
        {
            { 0x00000001, "CAN_COPY" },
            { 0x00000002, "CAN_MOVE" },
            { 0x00000004, "CAN_HAVESHORTCUT"},
            { 0x00000008, "STORAGE"},
            { 0x00000010, "CAN_RENAME" },
            { 0x00000020, "CAN_DELETE" },
            { 0x00000040, "HAS_PROPSHEET" },
            { 0x00000100, "IS_DROPTARGET" },
            { 0x00000177, "CAPABILITY_MASK" },
            { 0x00001000, "IS_SYSTEMITEM" },
            { 0x00002000, "IS_ENCRYPTED" },
            { 0x00004000, "IS_SLOW" },
            { 0x00008000, "IS_GHOSTED" },
            { 0x00010000, "IS_SHORTCUT" },
            { 0x00020000, "IS_SHARED" },
            { 0x00040000, "IS_READONLY" },
            { 0x00080000, "IS_HIDDEN" },
            { 0x000FC000, "DISPLAY_ATTRMASK" },
            { 0x00100000, "IS_NONENUMERATED" },
            { 0x00200000, "IS_NEWCONTENT" },
            { 0x00400000, "HAS_STREAM" },
            { 0x00800000, "IS_ANCESTOR" },
            { 0x01000000, "VALIDATE_ITEMS" },
            { 0x02000000, "REMOVABLE" },
            { 0x04000000, "IS_COMPRESSED" },
            { 0x08000000, "BROWSABLE" },
            { 0x10000000, "IS_FILE_SYS_ANCESTOR" },
            { 0x20000000, "IS_FOLDER" },
            { 0x40000000, "IS_FILESYSTEM_PART" },
            { 0x70C50008, "STORAGE_CAP_MASK" },
            { 0x80000000, "HAS_SUBFOLDER" },
        };

        private static readonly Dictionary<long, string> FileAttributes = new()
        /// https://learn.microsoft.com/en-us/windows/win32/fileio/file-attribute-constants
        {
            { 0x00000001, "READONLY" },
            { 0x00000002, "HIDDEN" },
            { 0x00000004, "SYSTEM" },
            { 0x00000010, "DIRECTORY" },
            { 0x00000020, "ARCHIVE" },
            { 0x00000040, "DEVICE" },
            { 0x00000080, "NORMAL" },
            { 0x00000100, "TEMPORARY" },
            { 0x00000200, "SPARSE" },
            { 0x00000400, "REPARSE_POINT" },
            { 0x00000800, "COMPRESSED" },
            { 0x00001000, "OFFLINE" },
            { 0x00002000, "NOT_CONTENT_INDEXED" },
            { 0x00004000, "ENCRYPTED" },
            { 0x00008000, "INTEGRITY_STREAM" },
            { 0x00010000, "VIRTUAL" },
            { 0x00020000, "NO_SCRUB_DATA" },
            { 0x00040000, "EA_OR_RECALL_ON_OPEN" },
            { 0x00080000, "PINNED_LOCAL" },
            { 0x00100000, "UNPINNED_LOCAL" },
            { 0x00400000, "RECALL_ON_DATA_ACCESS" }
        };

        private static readonly Dictionary<long, string> FilePlaceholderStates = new()
        /// https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/ne-shobjidl_core-placeholder_states
        {
            { 0, "None" },
            { 1, "Marked For Offline Availability" },
            { 2, "Available" },
            { 4, "Accessible" },
            { 8, "CloudFile" }
        };

        private static readonly (ulong min, string name)[] Capacity =
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-capacity
        [
            (0, "Empty"),
            (1, "Tiny"),
            (17179869185, "Small"),
            (85899345921, "Medium"),
            (274877906945, "Large"),
            (549755813889, "Huge"),
            (1099511627777, "Gigantic")
        ];

        public static readonly Dictionary<string, Dictionary<int, string>> LookupValues = new(StringComparer.OrdinalIgnoreCase)
        {
            { "System.Image.Compression", ImageCompression },
            { "System.SyncTransferStatusFlags", SyncTransferStatus },
            { "System.Importance", Importance },
            { "System.FlagColor", FlagColor },
            { "System.Calendar.ResponseStatus", ResponseStatus },
            { "System.Calendar.ShowTimeAs", ShowTimeAs },
            { "System.Sensitivity", Sensitivity },
            { "System.FlagStatus", FlagStatus },
            { "System.Communication.TaskStatus", TaskStatus }
        };

        
        private static readonly (ulong min, string name)[] MediaDurations = 
        /// https://learn.microsoft.com/en-us/windows/win32/properties/props-system-media-duration
        [    
            (0, "Very Short (under 1 min)"),
            (600000000, "Short (1 - 5 mins)"),
            (3000000000, "Medium (5 - 30 mins)"),
            (18000000000, "Long (30 - 60 mins)"),
            (36000000000, "Very Long (over 60 mins)"),
        ];

        private static string GetMeasurementName(ulong num, (ulong min, string name)[] list)
        {
            string result = "Unknown";
            foreach (var (min, name) in list)
            {
                if (num >= min)
                    result = name;
                else
                    break;
            }
            return result;
        }

        private static readonly Dictionary<string, (string guid, int propertyId)> PropertyKeys = new()
        /// PropertyKeys extracted from propkey.h from the Windows Kit
        /// Actual values can be found here: https://learn.microsoft.com/en-us/windows/win32/properties/props
        /// https://learn.microsoft.com/en-us/windows/win32/search/-search-3x-wds-propertymappings
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wsp/67328dcc-4e12-4e1e-be80-d91684df2f98
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wsp/2dbe759c-c955-4770-a545-e46d7f6332ed
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/bf7aeae8-c47a-4939-9f45-700158dac3bc
        /// Activity_ and ActivityHistory_ values from: https://cdn.callback.com/shellboost/doc/Reference/ShellBoost-Core-Assembly/html/f2805716-6023-897c-2234-d79a4feb5c5f.htm
        /// and also correlated between databases and Properties.
        /// Media Formats: https://gix.github.io/media-types/
        {
            { "Activity_AccountID", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 626) },
            { "Activity_ActivityId", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 620) },
            { "Activity_ActivationUri", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 640) },
            { "Activity_AppDisplayName", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 600) },
            { "Activity_AppIdKind", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 619) },
            { "Activity_AppIdList", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 635) },
            { "Activity_AppImageUri", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 601) },
            { "Activity_AttributionName", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 623) },
            { "Activity_BackgroundColor", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 602) },
            { "Activity_ContentImageUri", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 603) },
            { "Activity_ContentUri", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 604) },
            { "Activity_ContentVisualPropertiesHash", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 624) },
            { "Activity_Description", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 605) },
            { "Activity_DisplayText", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 606) },
            { "Activity_FallbackUri", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 641) },
            { "Activity_HasAdaptiveContent", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 625) },
            { "Activity_SetCategory", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 639) },
            { "Activity_SetId", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 634) },
            { "ActivityHistory_AppId", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 612) },
            { "ActivityHistory_ActiveDays", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 608) },
            { "ActivityHistory_ActiveDuration", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 609) },
            { "ActivityHistory_AppActivityId", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 611) },
            { "ActivityHistory_DaysActive", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 627) },
            { "ActivityHistory_DeviceId", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 614) },
            { "ActivityHistory_DeviceMake", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 632) },
            { "ActivityHistory_DeviceModel", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 633) },
            { "ActivityHistory_DeviceName", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 613) },
            { "ActivityHistory_DeviceType", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 631) },
            { "ActivityHistory_EndTime", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 616) },
            { "ActivityHistory_HoursActive", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 628) },
            { "ActivityHistory_Id", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 617) },
            { "ActivityHistory_Importance", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 637) },
            { "ActivityHistory_IsHistoryAttributedToSetAnchor", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 638) },
            { "ActivityHistory_IsLocal", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 630) },
            { "ActivityHistory_LocalEndTime", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 622) },
            { "ActivityHistory_LocalStartTime", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 621) },
            { "ActivityHistory_LocationActivityId", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 629) },
            { "ActivityHistory_StartTime", ("{c5043536-932e-219e-5fb9-1c2807d7b03e}", 618) },
            { "Address_Country", ("{c07b4199-e1df-4493-b1e1-de5946fb58f8}", 100) },
            { "Address_CountryCode", ("{c07b4199-e1df-4493-b1e1-de5946fb58f8}", 101) },
            { "Address_Region", ("{c07b4199-e1df-4493-b1e1-de5946fb58f8}", 102) },
            { "Address_RegionCode", ("{c07b4199-e1df-4493-b1e1-de5946fb58f8}", 103) },
            { "Address_Town", ("{c07b4199-e1df-4493-b1e1-de5946fb58f8}", 104) },
            { "Audio_ChannelCount", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 7) },
            { "Audio_Compression", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 10) },
            { "Audio_EncodingBitrate", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 4) },
            { "Audio_Format", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 2) },
            { "Audio_IsVariableBitRate", ("{e6822fee-8c17-4d62-823c-8e9cfcbd1d5c}", 100) },
            { "Audio_PeakValue", ("{2579e5d0-1116-4084-bd9a-9b4f7cb4df5e}", 100) },
            { "Audio_SampleRate", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 5) },
            { "Audio_SampleSize", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 6) },
            { "Audio_StreamName", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 9) },
            { "Audio_StreamNumber", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 8) },
            { "Calendar_Duration", ("{293ca35a-09aa-4dd2-b180-1fe245728a52}", 100) },
            { "Calendar_IsOnline", ("{bfee9149-e3e2-49a7-a862-c05988145cec}", 100) },
            { "Calendar_IsRecurring", ("{315b9c8d-80a9-4ef9-ae16-8e746da51d70}", 100) },
            { "Calendar_Location", ("{f6272d18-cecc-40b1-b26a-3911717aa7bd}", 100) },
            { "Calendar_OptionalAttendeeAddresses", ("{d55bae5a-3892-417a-a649-c6ac5aaaeab3}", 100) },
            { "Calendar_OptionalAttendeeNames", ("{09429607-582d-437f-84c3-de93a2b24c3c}", 100) },
            { "Calendar_OrganizerAddress", ("{744c8242-4df5-456c-ab9e-014efb9021e3}", 100) },
            { "Calendar_OrganizerName", ("{aaa660f9-9865-458e-b484-01bc7fe3973e}", 100) },
            { "Calendar_ReminderTime", ("{72fc5ba4-24f9-4011-9f3f-add27afad818}", 100) },
            { "Calendar_RequiredAttendeeAddresses", ("{0ba7d6c3-568d-4159-ab91-781a91fb71e5}", 100) },
            { "Calendar_RequiredAttendeeNames", ("{b33af30b-f552-4584-936c-cb93e5cda29f}", 100) },
            { "Calendar_Resources", ("{00f58a38-c54b-4c40-8696-97235980eae1}", 100) },
            { "Calendar_ResponseStatus", ("{188c1f91-3c40-4132-9ec5-d8b03b72a8a2}", 100) },
            { "Calendar_ShowTimeAs", ("{5bf396d4-5eb2-466f-bde9-2fb3f2361d6e}", 100) },
            { "Calendar_ShowTimeAsText", ("{53da57cf-62c0-45c4-81de-7610bcefd7f5}", 100) },
            { "Communication_AccountName", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 9) },
            { "Communication_DateItemExpires", ("{428040ac-a177-4c8a-9760-f6f761227f9a}", 100) },
            { "Communication_Direction", ("{8e531030-b960-4346-ae0d-66bc9a86fb94}", 100) },
            { "Communication_FollowupIconIndex", ("{83a6347e-6fe4-4f40-ba9c-c4865240d1f4}", 100) },
            { "Communication_HeaderItem", ("{c9c34f84-2241-4401-b607-bd20ed75ae7f}", 100) },
            { "Communication_PolicyTag", ("{ec0b4191-ab0b-4c66-90b6-c6637cdebbab}", 100) },
            { "Communication_SecurityFlags", ("{8619a4b6-9f4d-4429-8c0f-b996ca59e335}", 100) },
            { "Communication_Suffix", ("{807b653a-9e91-43ef-8f97-11ce04ee20c5}", 100) },
            { "Communication_TaskStatus", ("{be1a72c6-9a1d-46b7-afe7-afaf8cef4999}", 100) },
            { "Communication_TaskStatusText", ("{a6744477-c237-475b-a075-54f34498292a}", 100) },
            { "Computer_DecoratedFreeSpace", ("{9b174b35-40ff-11d2-a27e-00c04fc30871}", 7) },
            { "ConnectedSearch_ActivationCommand", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 18) },
            { "ConnectedSearch_AddOpenInBrowserCommand", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 29) },
            { "ConnectedSearch_AppInstalledState", ("{d76e7ba8-dfa6-48e7-9670-d62dfb07206b}", 4) },
            { "ConnectedSearch_ApplicationSearchScope", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 9) },
            { "ConnectedSearch_AppMinVersion", ("{d76e7ba8-dfa6-48e7-9670-d62dfb07206b}", 3) },
            { "ConnectedSearch_AutoComplete", ("{916d17ac-8a97-48af-85b7-867a88fad542}", 2) },
            { "ConnectedSearch_BypassViewAction", ("{dce33a78-aa18-4b3d-b1df-a6621ac8bdd2}", 2) },
            { "ConnectedSearch_ChildCount", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 11) },
            { "ConnectedSearch_ContractId", ("{d76e7ba8-dfa6-48e7-9670-d62dfb07206b}", 2) },
            { "ConnectedSearch_CopyText", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 28) },
            { "ConnectedSearch_DeferImagePrefetch", ("{a40294ef-d2b1-40ed-9512-dd3853b431f5}", 2) },
            { "ConnectedSearch_DisambiguationId", ("{f27abe3a-7111-4dda-8cb2-29222ae23566}", 2) },
            { "ConnectedSearch_DisambiguationText", ("{05e932b1-7ca2-491f-bd69-99b4cb266cbb}", 2) },
            { "ConnectedSearch_FallbackTemplate", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 4) },
            { "ConnectedSearch_HistoryDescription", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 22) },
            { "ConnectedSearch_HistoryGlyph", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 23) },
            { "ConnectedSearch_HistoryTitle", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 21) },
            { "ConnectedSearch_ImagePrefetchStage", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 31) },
            { "ConnectedSearch_ImageUri", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 30) },
            { "ConnectedSearch_ImpressionId", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 6) },
            { "ConnectedSearch_IsActivatable", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 141) },
            { "ConnectedSearch_IsAppAvailable", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 20) },
            { "ConnectedSearch_IsHistoryItem", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 19) },
            { "ConnectedSearch_IsLocalItem", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 32) },
            { "ConnectedSearch_IsVisibilityTracked", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 7) },
            { "ConnectedSearch_IsVisibleByDefault", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 13) },
            { "ConnectedSearch_ItemSource", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 17) },
            { "ConnectedSearch_JumpList", ("{e2d40928-632c-4280-a202-e0c2ad1ea0f4}", 3) },
            { "ConnectedSearch_LinkText", ("{12fa14f5-c6fe-4545-bce2-1ed6cb6b8422}", 2) },
            { "ConnectedSearch_LocalWeights", ("{79486778-4c6f-4dde-bc53-cd594311af99}", 2) },
            { "ConnectedSearch_ParentId", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 10) },
            { "ConnectedSearch_QsCode", ("{e2d40928-632c-4280-a202-e0c2ad1ea0f4}", 2) },
            { "ConnectedSearch_ReferrerId", ("{a8a7a412-1927-4a34-b1d4-45f67cc672fb}", 2) },
            { "ConnectedSearch_RegionId", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 16) },
            { "ConnectedSearch_RenderingTemplate", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 3) },
            { "ConnectedSearch_RequireInstall", ("{cc158e89-6581-4311-9637-a8da9002f118}", 2) },
            { "ConnectedSearch_RequiresConsent", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 27) },
            { "ConnectedSearch_RequireTemplate", ("{73389854-0b42-4ea6-bc67-847d430899fd}", 2) },
            { "ConnectedSearch_SuggestionContext", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 15) },
            { "ConnectedSearch_SuggestionDetailText", ("{e9641eff-af25-4db7-947b-4128929f8ef5}", 2) },
            { "ConnectedSearch_SuppressLocalHero", ("{b769d0fe-bc33-421a-8ce6-45add82ec756}", 2) },
            { "ConnectedSearch_TelemetryData", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 8) },
            { "ConnectedSearch_TelemetryId", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 5) },
            { "ConnectedSearch_TopLevelId", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 12) },
            { "ConnectedSearch_Type", ("{e1ad4953-a752-443c-93bf-80c7525566c2}", 2) },
            { "ConnectedSearch_VoiceCommandExamples", ("{e2d40928-632c-4280-a202-e0c2ad1ea0f4}", 4) },
            { "Contact_AccountPictureDynamicVideo", ("{0b8bb018-2725-4b44-92ba-7933aeb2dde7}", 2) },
            { "Contact_AccountPictureLarge", ("{0b8bb018-2725-4b44-92ba-7933aeb2dde7}", 3) },
            { "Contact_AccountPictureSmall", ("{0b8bb018-2725-4b44-92ba-7933aeb2dde7}", 4) },
            { "Contact_Anniversary", ("{9ad5badb-cea7-4470-a03d-b84e51b9949e}", 100) },
            { "Contact_AssistantName", ("{cd102c9c-5540-4a88-a6f6-64e4981c8cd1}", 100) },
            { "Contact_AssistantTelephone", ("{9a93244d-a7ad-4ff8-9b99-45ee4cc09af6}", 100) },
            { "Contact_Birthday", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 47) },
            { "Contact_BusinessAddress", ("{730fb6dd-cf7c-426b-a03f-bd166cc9ee24}", 100) },
            { "Contact_BusinessAddress1Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 119) },
            { "Contact_BusinessAddress1Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 117) },
            { "Contact_BusinessAddress1PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 120) },
            { "Contact_BusinessAddress1Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 118) },
            { "Contact_BusinessAddress1Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 116) },
            { "Contact_BusinessAddress2Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 124) },
            { "Contact_BusinessAddress2Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 122) },
            { "Contact_BusinessAddress2PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 125) },
            { "Contact_BusinessAddress2Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 123) },
            { "Contact_BusinessAddress2Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 121) },
            { "Contact_BusinessAddress3Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 129) },
            { "Contact_BusinessAddress3Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 127) },
            { "Contact_BusinessAddress3PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 130) },
            { "Contact_BusinessAddress3Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 128) },
            { "Contact_BusinessAddress3Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 126) },
            { "Contact_BusinessAddressCity", ("{402b5934-ec5a-48c3-93e6-85e86a2d934e}", 100) },
            { "Contact_BusinessAddressCountry", ("{b0b87314-fcf6-4feb-8dff-a50da6af561c}", 100) },
            { "Contact_BusinessAddressPostalCode", ("{e1d4a09e-d758-4cd1-b6ec-34a8b5a73f80}", 100) },
            { "Contact_BusinessAddressPostOfficeBox", ("{bc4e71ce-17f9-48d5-bee9-021df0ea5409}", 100) },
            { "Contact_BusinessAddressState", ("{446f787f-10c4-41cb-a6c4-4d0343551597}", 100) },
            { "Contact_BusinessAddressStreet", ("{ddd1460f-c0bf-4553-8ce4-10433c908fb0}", 100) },
            { "Contact_BusinessEmailAddresses", ("{f271c659-7e5e-471f-ba25-7f77b286f836}", 100) },
            { "Contact_BusinessFaxNumber", ("{91eff6f3-2e27-42ca-933e-7c999fbe310b}", 100) },
            { "Contact_BusinessHomePage", ("{56310920-2491-4919-99ce-eadb06fafdb2}", 100) },
            { "Contact_BusinessTelephone", ("{6a15e5a0-0a1e-4cd7-bb8c-d2f1b0c929bc}", 100) },
            { "Contact_CallbackTelephone", ("{bf53d1c3-49e0-4f7f-8567-5a821d8ac542}", 100) },
            { "Contact_CarTelephone", ("{8fdc6dea-b929-412b-ba90-397a257465fe}", 100) },
            { "Contact_Children", ("{d4729704-8ef1-43ef-9024-2bd381187fd5}", 100) },
            { "Contact_CompanyMainTelephone", ("{8589e481-6040-473d-b171-7fa89c2708ed}", 100) },
            { "Contact_ConnectedServiceDisplayName", ("{39b77f4f-a104-4863-b395-2db2ad8f7bc1}", 100) },
            { "Contact_ConnectedServiceIdentities", ("{80f41eb8-afc4-4208-aa5f-cce21a627281}", 100) },
            { "Contact_ConnectedServiceName", ("{b5c84c9e-5927-46b5-a3cc-933c21b78469}", 100) },
            { "Contact_ConnectedServiceSupportedActions", ("{a19fb7a9-024b-4371-a8bf-4d29c3e4e9c9}", 100) },
            { "Contact_DataSuppliers", ("{9660c283-fc3a-4a08-a096-eed3aac46da2}", 100) },
            { "Contact_Department", ("{fc9f7306-ff8f-4d49-9fb6-3ffe5c0951ec}", 100) },
            { "Contact_DisplayBusinessPhoneNumbers", ("{364028da-d895-41fe-a584-302b1bb70a76}", 100) },
            { "Contact_DisplayHomePhoneNumbers", ("{5068bcdf-d697-4d85-8c53-1f1cdab01763}", 100) },
            { "Contact_DisplayMobilePhoneNumbers", ("{9cb0c358-9d7a-46b1-b466-dcc6f1a3d93d}", 100) },
            { "Contact_DisplayOtherPhoneNumbers", ("{03089873-8ee8-4191-bd60-d31f72b7900b}", 100) },
            { "Contact_EmailAddress", ("{f8fa7fa3-d12b-4785-8a4e-691a94f7a3e7}", 100) },
            { "Contact_EmailAddress2", ("{38965063-edc8-4268-8491-b7723172cf29}", 100) },
            { "Contact_EmailAddress3", ("{644d37b4-e1b3-4bad-b099-7e7c04966aca}", 100) },
            { "Contact_EmailAddresses", ("{84d8f337-981d-44b3-9615-c7596dba17e3}", 100) },
            { "Contact_EmailName", ("{cc6f4f24-6083-4bd4-8754-674d0de87ab8}", 100) },
            { "Contact_FileAsName", ("{f1a24aa7-9ca7-40f6-89ec-97def9ffe8db}", 100) },
            { "Contact_FirstName", ("{14977844-6b49-4aad-a714-a4513bf60460}", 100) },
            { "Contact_FullName", ("{635e9051-50a5-4ba2-b9db-4ed056c77296}", 100) },
            { "Contact_Gender", ("{3c8cee58-d4f0-4cf9-b756-4e5d24447bcd}", 100) },
            { "Contact_GenderValue", ("{3c8cee58-d4f0-4cf9-b756-4e5d24447bcd}", 101) },
            { "Contact_Hobbies", ("{5dc2253f-5e11-4adf-9cfe-910dd01e3e70}", 100) },
            { "Contact_HomeAddress", ("{98f98354-617a-46b8-8560-5b1b64bf1f89}", 100) },
            { "Contact_HomeAddress1Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 104) },
            { "Contact_HomeAddress1Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 102) },
            { "Contact_HomeAddress1PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 105) },
            { "Contact_HomeAddress1Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 103) },
            { "Contact_HomeAddress1Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 101) },
            { "Contact_HomeAddress2Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 109) },
            { "Contact_HomeAddress2Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 107) },
            { "Contact_HomeAddress2PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 110) },
            { "Contact_HomeAddress2Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 108) },
            { "Contact_HomeAddress2Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 106) },
            { "Contact_HomeAddress3Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 114) },
            { "Contact_HomeAddress3Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 112) },
            { "Contact_HomeAddress3PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 115) },
            { "Contact_HomeAddress3Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 113) },
            { "Contact_HomeAddress3Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 111) },
            { "Contact_HomeAddressCity", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 65) },
            { "Contact_HomeAddressCountry", ("{08a65aa1-f4c9-43dd-9ddf-a33d8e7ead85}", 100) },
            { "Contact_HomeAddressPostalCode", ("{8afcc170-8a46-4b53-9eee-90bae7151e62}", 100) },
            { "Contact_HomeAddressPostOfficeBox", ("{7b9f6399-0a3f-4b12-89bd-4adc51c918af}", 100) },
            { "Contact_HomeAddressState", ("{c89a23d0-7d6d-4eb8-87d4-776a82d493e5}", 100) },
            { "Contact_HomeAddressStreet", ("{0adef160-db3f-4308-9a21-06237b16fa2a}", 100) },
            { "Contact_HomeEmailAddresses", ("{56c90e9d-9d46-4963-886f-2e1cd9a694ef}", 100) },
            { "Contact_HomeFaxNumber", ("{660e04d6-81ab-4977-a09f-82313113ab26}", 100) },
            { "Contact_HomeTelephone", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 20) },
            { "Contact_IMAddress", ("{d68dbd8a-3374-4b81-9972-3ec30682db3d}", 100) },
            { "Contact_Initials", ("{f3d8f40d-50cb-44a2-9718-40cb9119495d}", 100) },
            { "Contact_JA_CompanyNamePhonetic", ("{897b3694-fe9e-43e6-8066-260f590c0100}", 2) },
            { "Contact_JA_FirstNamePhonetic", ("{897b3694-fe9e-43e6-8066-260f590c0100}", 3) },
            { "Contact_JA_LastNamePhonetic", ("{897b3694-fe9e-43e6-8066-260f590c0100}", 4) },
            { "Contact_JobInfo1CompanyAddress", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 120) },
            { "Contact_JobInfo1CompanyName", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 102) },
            { "Contact_JobInfo1Department", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 106) },
            { "Contact_JobInfo1Manager", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 105) },
            { "Contact_JobInfo1OfficeLocation", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 104) },
            { "Contact_JobInfo1Title", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 103) },
            { "Contact_JobInfo1YomiCompanyName", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 101) },
            { "Contact_JobInfo2CompanyAddress", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 121) },
            { "Contact_JobInfo2CompanyName", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 108) },
            { "Contact_JobInfo2Department", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 113) },
            { "Contact_JobInfo2Manager", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 112) },
            { "Contact_JobInfo2OfficeLocation", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 110) },
            { "Contact_JobInfo2Title", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 109) },
            { "Contact_JobInfo2YomiCompanyName", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 107) },
            { "Contact_JobInfo3CompanyAddress", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 123) },
            { "Contact_JobInfo3CompanyName", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 115) },
            { "Contact_JobInfo3Department", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 119) },
            { "Contact_JobInfo3Manager", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 118) },
            { "Contact_JobInfo3OfficeLocation", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 117) },
            { "Contact_JobInfo3Title", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 116) },
            { "Contact_JobInfo3YomiCompanyName", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 114) },
            { "Contact_JobTitle", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 6) },
            { "Contact_Label", ("{97b0ad89-df49-49cc-834e-660974fd755b}", 100) },
            { "Contact_LastName", ("{8f367200-c270-457c-b1d4-e07c5bcd90c7}", 100) },
            { "Contact_MailingAddress", ("{c0ac206a-827e-4650-95ae-77e2bb74fcc9}", 100) },
            { "Contact_MiddleName", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 71) },
            { "Contact_MobileTelephone", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 35) },
            { "Contact_NickName", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 74) },
            { "Contact_OfficeLocation", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 7) },
            { "Contact_OtherAddress", ("{508161fa-313b-43d5-83a1-c1accf68622c}", 100) },
            { "Contact_OtherAddress1Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 134) },
            { "Contact_OtherAddress1Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 132) },
            { "Contact_OtherAddress1PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 135) },
            { "Contact_OtherAddress1Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 133) },
            { "Contact_OtherAddress1Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 131) },
            { "Contact_OtherAddress2Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 139) },
            { "Contact_OtherAddress2Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 137) },
            { "Contact_OtherAddress2PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 140) },
            { "Contact_OtherAddress2Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 138) },
            { "Contact_OtherAddress2Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 136) },
            { "Contact_OtherAddress3Country", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 144) },
            { "Contact_OtherAddress3Locality", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 142) },
            { "Contact_OtherAddress3PostalCode", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 145) },
            { "Contact_OtherAddress3Region", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 143) },
            { "Contact_OtherAddress3Street", ("{a7b6f596-d678-4bc1-b05f-0203d27e8aa1}", 141) },
            { "Contact_OtherAddressCity", ("{6e682923-7f7b-4f0c-a337-cfca296687bf}", 100) },
            { "Contact_OtherAddressCountry", ("{8f167568-0aae-4322-8ed9-6055b7b0e398}", 100) },
            { "Contact_OtherAddressPostalCode", ("{95c656c1-2abf-4148-9ed3-9ec602e3b7cd}", 100) },
            { "Contact_OtherAddressPostOfficeBox", ("{8b26ea41-058f-43f6-aecc-4035681ce977}", 100) },
            { "Contact_OtherAddressState", ("{71b377d6-e570-425f-a170-809fae73e54e}", 100) },
            { "Contact_OtherAddressStreet", ("{ff962609-b7d6-4999-862d-95180d529aea}", 100) },
            { "Contact_OtherEmailAddresses", ("{11d6336b-38c4-4ec9-84d6-eb38d0b150af}", 100) },
            { "Contact_PagerTelephone", ("{d6304e01-f8f5-4f45-8b15-d024a6296789}", 100) },
            { "Contact_PersonalTitle", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 69) },
            { "Contact_PhoneNumbersCanonical", ("{d042d2a1-927e-40b5-a503-6edbd42a517e}", 100) },
            { "Contact_Prefix", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 75) },
            { "Contact_PrimaryAddressCity", ("{c8ea94f0-a9e3-4969-a94b-9c62a95324e0}", 100) },
            { "Contact_PrimaryAddressCountry", ("{e53d799d-0f3f-466e-b2ff-74634a3cb7a4}", 100) },
            { "Contact_PrimaryAddressPostalCode", ("{18bbd425-ecfd-46ef-b612-7b4a6034eda0}", 100) },
            { "Contact_PrimaryAddressPostOfficeBox", ("{de5ef3c7-46e1-484e-9999-62c5308394c1}", 100) },
            { "Contact_PrimaryAddressState", ("{f1176dfe-7138-4640-8b4c-ae375dc70a6d}", 100) },
            { "Contact_PrimaryAddressStreet", ("{63c25b20-96be-488f-8788-c09c407ad812}", 100) },
            { "Contact_PrimaryEmailAddress", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 48) },
            { "Contact_PrimaryTelephone", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 25) },
            { "Contact_Profession", ("{7268af55-1ce4-4f6e-a41f-b6e4ef10e4a9}", 100) },
            { "Contact_SpouseName", ("{9d2408b6-3167-422b-82b0-f583b7a7cfe3}", 100) },
            { "Contact_Suffix", ("{176dc63c-2688-4e89-8143-a347800f25e9}", 73) },
            { "Contact_TelexNumber", ("{c554493c-c1f7-40c1-a76c-ef8c0614003e}", 100) },
            { "Contact_TTYTDDTelephone", ("{aaf16bac-2b55-45e6-9f6d-415eb94910df}", 100) },
            { "Contact_WebPage", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 18) },
            { "Contact_Webpage2", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 124) },
            { "Contact_Webpage3", ("{00f63dd8-22bd-4a5d-ba34-5cb0b9bdcb03}", 125) },
            { "AcquisitionID", ("{65a98875-3c80-40ab-abbc-efdaf77dbee2}", 100) },
            { "ApplicationDefinedProperties", ("{cdbfc167-337e-41d8-af7c-8c09205429c7}", 100) },
            { "ApplicationName", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 18) },
            { "AppZoneIdentifier", ("{502cfeab-47eb-459c-b960-e6d8728f7701}", 102) },
            { "Author", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 4) },
            { "CachedFileUpdaterContentIdForConflictResolution", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 114) },
            { "CachedFileUpdaterContentIdForStream", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 113) },
            { "Capacity", ("{9b174b35-40ff-11d2-a27e-00c04fc30871}", 3) },
            { "Category", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 2) },
            { "Comment", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 6) },
            { "Company", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 15) },
            { "ComputerName", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 5) },
            { "ContainedItems", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 29) },
            { "ContentId", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 132) },
            { "ContentStatus", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 27) },
            { "ContentType", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 26) },
            { "ContentUri", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 131) },
            { "Copyright", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 11) },
            { "CreatorAppId", ("{c2ea046e-033c-4e91-bd5b-d4942f6bbe49}", 2) },
            { "CreatorOpenWithUIOptions", ("{c2ea046e-033c-4e91-bd5b-d4942f6bbe49}", 3) },
            { "DataObjectFormat", ("{1e81a3f8-a30f-4247-b9ee-1d0368a9425c}", 2) },
            { "DateAccessed", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 16) },
            { "DateAcquired", ("{2cbaa8f5-d81f-47ca-b17a-f8d822300131}", 100) },
            { "DateArchived", ("{43f8d7b7-a444-4f87-9383-52271c9b915c}", 100) },
            { "DateCompleted", ("{72fab781-acda-43e5-b155-b2434f85e678}", 100) },
            { "DateCreated", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 15) },
            { "DateImported", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 18258) },
            { "DateModified", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 14) },
            { "DefaultSaveLocationDisplay", ("{5d76b67f-9b3d-44bb-b6ae-25da4f638a67}", 10) },
            { "DueDate", ("{3f8472b5-e0af-4db2-8071-c53fe76ae7ce}", 100) },
            { "EndDate", ("{c75faa05-96fd-49e7-9cb4-9f601082d553}", 100) },
            { "ExpandoProperties", ("{6fa20de6-d11c-4d9d-a154-64317628c12d}", 100) },
            { "FileAllocationSize", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 18) },
            { "FileAttributes", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 13) },
            { "FileCount", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 12) },
            { "FileDescription", ("{0cef7d53-fa64-11d1-a203-0000f81fedee}", 3) },
            { "FileExtension", ("{e4f10a3c-49e6-405d-8288-a23bd4eeaa6c}", 100) },
            { "FileFRN", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 21) },
            { "FileName", ("{41cf5ae0-f75a-4806-bd87-59c7d9248eb9}", 100) },
            { "FileOfflineAvailabilityStatus", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 100) },
            { "FileOwner", ("{9b174b34-40ff-11d2-a27e-00c04fc30871}", 4) },
            { "FilePlaceholderStatus", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 2) },
            { "FileVersion", ("{0cef7d53-fa64-11d1-a203-0000f81fedee}", 4) },
            { "FindData", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 0) },
            { "FlagColor", ("{67df94de-0ca7-4d6f-b792-053a3e4f03cf}", 100) },
            { "FlagColorText", ("{45eae747-8e2a-40ae-8cbf-ca52aba6152a}", 100) },
            { "FlagStatus", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 12) },
            { "FlagStatusText", ("{dc54fd2e-189d-4871-aa01-08c2f57a4abc}", 100) },
            { "FolderKind", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 101) },
            { "FolderNameDisplay", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 25) },
            { "FreeSpace", ("{9b174b35-40ff-11d2-a27e-00c04fc30871}", 2) },
            { "FullText", ("{1e3ee840-bc2b-476c-8237-2acd1a839b22}", 6) },
            { "HighKeywords", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 24) },
            { "Identity", ("{a26f4afc-7346-4299-be47-eb1ae613139f}", 100) },
            { "Identity_Blob", ("{8c3b93a4-baed-1a83-9a32-102ee313f6eb}", 100) },
            { "Identity_DisplayName", ("{7d683fc9-d155-45a8-bb1f-89d19bcb792f}", 100) },
            { "Identity_InternetSid", ("{6d6d5d49-265d-4688-9f4e-1fdd33e7cc83}", 100) },
            { "Identity_IsMeIdentity", ("{a4108708-09df-4377-9dfc-6d99986d5a67}", 100) },
            { "Identity_KeyProviderContext", ("{a26f4afc-7346-4299-be47-eb1ae613139f}", 17) },
            { "Identity_KeyProviderName", ("{a26f4afc-7346-4299-be47-eb1ae613139f}", 16) },
            { "Identity_LogonStatusString", ("{f18dedf3-337f-42c0-9e03-cee08708a8c3}", 100) },
            { "Identity_PrimaryEmailAddress", ("{fcc16823-baed-4f24-9b32-a0982117f7fa}", 100) },
            { "Identity_PrimarySid", ("{2b1b801e-c0c1-4987-9ec5-72fa89814787}", 100) },
            { "Identity_ProviderData", ("{a8a74b92-361b-4e9a-b722-7c4a7330a312}", 100) },
            { "Identity_ProviderID", ("{74a7de49-fa11-4d3d-a006-db7e08675916}", 100) },
            { "Identity_QualifiedUserName", ("{da520e51-f4e9-4739-ac82-02e0a95c9030}", 100) },
            { "Identity_UniqueID", ("{e55fc3b0-2b60-4220-918e-b21e8bf16016}", 100) },
            { "Identity_UserName", ("{c4322503-78ca-49c6-9acc-a68e2afd7b6b}", 100) },
            { "IdentityProvider_Name", ("{b96eff7b-35ca-4a35-8607-29e3a54c46ea}", 100) },
            { "IdentityProvider_Picture", ("{2425166f-5642-4864-992f-98fd98f294c3}", 100) },
            { "ImageParsingName", ("{d7750ee0-c6a4-48ec-b53e-b87b52e6d073}", 100) },
            { "Importance", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 11) },
            { "ImportanceText", ("{a3b29791-7713-4e1d-bb40-17db85f01831}", 100) },
            { "IsAttachment", ("{f23f425c-71a1-4fa8-922f-678ea4a60408}", 100) },
            { "IsDefaultNonOwnerSaveLocation", ("{5d76b67f-9b3d-44bb-b6ae-25da4f638a67}", 5) },
            { "IsDefaultSaveLocation", ("{5d76b67f-9b3d-44bb-b6ae-25da4f638a67}", 3) },
            { "IsDeleted", ("{5cda5fc8-33ee-4ff3-9094-ae7bd8868c4d}", 100) },
            { "IsEncrypted", ("{90e5e14e-648b-4826-b2aa-acaf790e3513}", 10) },
            { "IsFlagged", ("{5da84765-e3ff-4278-86b0-a27967fbdd03}", 100) },
            { "IsFlaggedComplete", ("{a6f360d2-55f9-48de-b909-620e090a647c}", 100) },
            { "IsIncomplete", ("{346c8bd1-2e6a-4c45-89a4-61b78e8e700f}", 100) },
            { "IsLocationSupported", ("{5d76b67f-9b3d-44bb-b6ae-25da4f638a67}", 8) },
            { "IsPinnedToNameSpaceTree", ("{5d76b67f-9b3d-44bb-b6ae-25da4f638a67}", 2) },
            { "IsRead", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 10) },
            { "IsSearchOnlyItem", ("{5d76b67f-9b3d-44bb-b6ae-25da4f638a67}", 4) },
            { "IsSendToTarget", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 33) },
            { "IsShared", ("{ef884c5b-2bfe-41bb-aae5-76eedf4f9902}", 100) },
            { "ItemAuthors", ("{d0a04f0a-462a-48a4-bb2f-3706e88dbd7d}", 100) },
            { "ItemClassType", ("{048658ad-2db8-41a4-bbb6-ac1ef1207eb1}", 100) },
            { "ItemDate", ("{f7db74b4-4287-4103-afba-f1b13dcd75cf}", 100) },
            { "ItemFolderNameDisplay", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 2) },
            { "ItemFolderPathDisplay", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 6) },
            { "ItemFolderPathDisplayNarrow", ("{dabd30ed-0043-4789-a7f8-d013a4736622}", 100) },
            { "ItemName", ("{6b8da074-3b5c-43bc-886f-0a2cdce00b6f}", 100) },
            { "ItemNameDisplay", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 10) },
            { "ItemNameDisplayWithoutExtension", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 24) },
            { "ItemNamePrefix", ("{d7313ff1-a77a-401c-8c99-3dbdd68add36}", 100) },
            { "ItemNameSortOverride", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 23) },
            { "ItemParticipants", ("{d4d0aa16-9948-41a4-aa85-d97ff9646993}", 100) },
            { "ItemPathDisplay", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 7) },
            { "ItemPathDisplayNarrow", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 8) },
            { "ItemSubType", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 37) },
            { "ItemType", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 11) },
            { "ItemTypeText", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 4) },
            { "ItemUrl", ("{49691c90-7e17-101a-a91c-08002b2ecda9}", 9) },
            { "Keywords", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 5) },
            { "Kind", ("{1e3ee840-bc2b-476c-8237-2acd1a839b22}", 3) },
            { "KindText", ("{f04bef95-c585-4197-a2b7-df46fdc9ee6d}", 100) },
            { "Language", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 28) },
            { "LastSyncError", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 107) },
            { "LastSyncWarning", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 128) },
            { "LastWriterPackageFamilyName", ("{502cfeab-47eb-459c-b960-e6d8728f7701}", 101) },
            { "LowKeywords", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 25) },
            { "MediumKeywords", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 26) },
            { "MileageInformation", ("{fdf84370-031a-4add-9e91-0d775f1c6605}", 100) },
            { "MIMEType", ("{0b63e350-9ccc-11d0-bcdb-00805fccce04}", 5) },
            { "Null", ("{00000000-0000-0000-0000-000000000000}", 0) },
            { "OfflineAvailability", ("{a94688b6-7d9f-4570-a648-e3dfc0ab2b3f}", 100) },
            { "OfflineStatus", ("{6d24888f-4718-4bda-afed-ea0fb4386cd8}", 100) },
            { "OriginalFileName", ("{0cef7d53-fa64-11d1-a203-0000f81fedee}", 6) },
            { "OwnerSID", ("{5d76b67f-9b3d-44bb-b6ae-25da4f638a67}", 6) },
            { "ParentalRating", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 21) },
            { "ParentalRatingReason", ("{10984e0a-f9f2-4321-b7ef-baf195af4319}", 100) },
            { "ParentalRatingsOrganization", ("{a7fe0840-1344-46f0-8d37-52ed712a4bf9}", 100) },
            { "ParsingBindContext", ("{dfb9a04d-362f-4ca3-b30b-0254b17b5b84}", 100) },
            { "ParsingName", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 24) },
            { "ParsingPath", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 30) },
            { "PerceivedType", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 9) },
            { "PercentFull", ("{9b174b35-40ff-11d2-a27e-00c04fc30871}", 5) },
            { "Priority", ("{9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4}", 5) },
            { "PriorityText", ("{d98be98b-b86b-4095-bf52-9d23b2e0a752}", 100) },
            { "Project", ("{39a7f922-477c-48de-8bc8-b28441e342e3}", 100) },
            { "ProviderItemID", ("{f21d9941-81f0-471a-adee-4e74b49217ed}", 100) },
            { "Rating", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 9) },
            { "RatingText", ("{90197ca7-fd8f-4e8c-9da3-b57e1e609295}", 100) },
            { "RemoteConflictingFile", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 115) },
            { "Security_AllowedEnterpriseDataProtectionIdentities", ("{38d43380-d418-4830-84d5-46935a81c5c6}", 32) },
            { "Security_EncryptionOwners", ("{5f5aff6a-37e5-4780-97ea-80c7565cf535}", 34) },
            { "Security_EncryptionOwnersDisplay", ("{de621b8f-e125-43a3-a32d-5665446d632a}", 25) },
            { "Sensitivity", ("{f8d3f6ac-4874-42cb-be59-ab454b30716a}", 100) },
            { "SensitivityText", ("{d0c7f054-3f72-4725-8527-129a577cb269}", 100) },
            { "SFGAOFlags", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 25) },
            { "SharedWith", ("{ef884c5b-2bfe-41bb-aae5-76eedf4f9902}", 200) },
            { "ShareUserRating", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 12) },
            { "SharingStatus", ("{ef884c5b-2bfe-41bb-aae5-76eedf4f9902}", 300) },
            { "Shell_OmitFromView", ("{de35258c-c695-4cbc-b982-38b0ad24ced0}", 2) },
            { "SimpleRating", ("{a09f084e-ad41-489f-8076-aa5be3082bca}", 100) },
            { "Size", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 12) },
            { "SoftwareUsed", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 305) },
            { "SourceItem", ("{668cdfa5-7a1b-4323-ae4b-e527393a1d81}", 100) },
            { "SourcePackageFamilyName", ("{ffae9db7-1c8d-43ff-818c-84403aa3732d}", 100) },
            { "StartDate", ("{48fd6ec8-8a12-4cdf-a03e-4ec5a511edde}", 100) },
            { "Status", ("{000214a1-0000-0000-c000-000000000046}", 9) },
            { "StorageProviderCallerVersionInformation", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 7) },
            { "StorageProviderCustomPrimaryIcon", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 12) },
            { "StorageProviderError", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 109) },
            { "StorageProviderFileChecksum", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 5) },
            { "StorageProviderFileCreatedBy", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 10) },
            { "StorageProviderFileFlags", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 8) },
            { "StorageProviderFileHasConflict", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 9) },
            { "StorageProviderFileIdentifier", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 3) },
            { "StorageProviderFileModifiedBy", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 11) },
            { "StorageProviderFileRemoteUri", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 112) },
            { "StorageProviderFileVersion", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 4) },
            { "StorageProviderFileVersionWaterline", ("{b2f9b9d6-fec4-4dd5-94d7-8957488c807b}", 6) },
            { "StorageProviderId", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 108) },
            { "StorageProviderShareStatuses", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 111) },
            { "StorageProviderSharingStatus", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 117) },
            { "StorageProviderStatus", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 110) },
            { "Subject", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 3) },
            { "SyncTransferStatus", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 103) },
            { "Thumbnail", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 17) },
            { "ThumbnailCacheId", ("{446d16b1-8dad-4870-a748-402ea43d788c}", 100) },
            { "ThumbnailStream", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 27) },
            { "Title", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 2) },
            { "TitleSortOverride", ("{f0f7984d-222e-4ad2-82ab-1dd8ea40e57e}", 300) },
            { "TotalFileSize", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 14) },
            { "Trademarks", ("{0cef7d53-fa64-11d1-a203-0000f81fedee}", 9) },
            { "TransferOrder", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 106) },
            { "TransferPosition", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 104) },
            { "TransferSize", ("{fceff153-e839-4cf3-a9e7-ea22832094b8}", 105) },
            { "VolumeId", ("{446d16b1-8dad-4870-a748-402ea43d788c}", 104) },
            { "ZoneIdentifier", ("{502cfeab-47eb-459c-b960-e6d8728f7701}", 100) },
            { "Device_PrinterURL", ("{0b48f35a-be6e-4f17-b108-3c4073d1669a}", 15) },
            { "DeviceInterface_Bluetooth_DeviceAddress", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 1) },
            { "DeviceInterface_Bluetooth_Flags", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 3) },
            { "DeviceInterface_Bluetooth_LastConnectedTime", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 11) },
            { "DeviceInterface_Bluetooth_Manufacturer", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 4) },
            { "DeviceInterface_Bluetooth_ModelNumber", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 5) },
            { "DeviceInterface_Bluetooth_ProductId", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 8) },
            { "DeviceInterface_Bluetooth_ProductVersion", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 9) },
            { "DeviceInterface_Bluetooth_ServiceGuid", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 2) },
            { "DeviceInterface_Bluetooth_VendorId", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 7) },
            { "DeviceInterface_Bluetooth_VendorIdSource", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 6) },
            { "DeviceInterface_Hid_IsReadOnly", ("{cbf38310-4a17-4310-a1eb-247f0b67593b}", 4) },
            { "DeviceInterface_Hid_ProductId", ("{cbf38310-4a17-4310-a1eb-247f0b67593b}", 6) },
            { "DeviceInterface_Hid_UsageId", ("{cbf38310-4a17-4310-a1eb-247f0b67593b}", 3) },
            { "DeviceInterface_Hid_UsagePage", ("{cbf38310-4a17-4310-a1eb-247f0b67593b}", 2) },
            { "DeviceInterface_Hid_VendorId", ("{cbf38310-4a17-4310-a1eb-247f0b67593b}", 5) },
            { "DeviceInterface_Hid_VersionNumber", ("{cbf38310-4a17-4310-a1eb-247f0b67593b}", 7) },
            { "DeviceInterface_PrinterDriverDirectory", ("{847c66de-b8d6-4af9-abc3-6f4f926bc039}", 14) },
            { "DeviceInterface_PrinterDriverName", ("{afc47170-14f5-498c-8f30-b0d19be449c6}", 11) },
            { "DeviceInterface_PrinterEnumerationFlag", ("{a00742a1-cd8c-4b37-95ab-70755587767a}", 3) },
            { "DeviceInterface_PrinterName", ("{0a7b84ef-0c27-463f-84ef-06c5070001be}", 10) },
            { "DeviceInterface_PrinterPortName", ("{eec7b761-6f94-41b1-949f-c729720dd13c}", 12) },
            { "DeviceInterface_Proximity_SupportsNfc", ("{fb3842cd-9e2a-4f83-8fcc-4b0761139ae9}", 2) },
            { "DeviceInterface_Serial_PortName", ("{4c6bf15c-4c03-4aac-91f5-64c0f852bcf4}", 4) },
            { "DeviceInterface_Serial_UsbProductId", ("{4c6bf15c-4c03-4aac-91f5-64c0f852bcf4}", 3) },
            { "DeviceInterface_Serial_UsbVendorId", ("{4c6bf15c-4c03-4aac-91f5-64c0f852bcf4}", 2) },
            { "DeviceInterface_WinUsb_DeviceInterfaceClasses", ("{95e127b5-79cc-4e83-9c9e-8422187b3e0e}", 7) },
            { "DeviceInterface_WinUsb_UsbClass", ("{95e127b5-79cc-4e83-9c9e-8422187b3e0e}", 4) },
            { "DeviceInterface_WinUsb_UsbProductId", ("{95e127b5-79cc-4e83-9c9e-8422187b3e0e}", 3) },
            { "DeviceInterface_WinUsb_UsbProtocol", ("{95e127b5-79cc-4e83-9c9e-8422187b3e0e}", 6) },
            { "DeviceInterface_WinUsb_UsbSubClass", ("{95e127b5-79cc-4e83-9c9e-8422187b3e0e}", 5) },
            { "DeviceInterface_WinUsb_UsbVendorId", ("{95e127b5-79cc-4e83-9c9e-8422187b3e0e}", 2) },
            { "Devices_Aep_AepId", ("{3b2ce006-5e61-4fde-bab8-9b8aac9b26df}", 8) },
            { "Devices_Aep_Bluetooth_Cod_Major", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 2) },
            { "Devices_Aep_Bluetooth_Cod_Minor", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 3) },
            { "Devices_Aep_Bluetooth_Cod_Services_Audio", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 10) },
            { "Devices_Aep_Bluetooth_Cod_Services_Capturing", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 8) },
            { "Devices_Aep_Bluetooth_Cod_Services_Information", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 12) },
            { "Devices_Aep_Bluetooth_Cod_Services_LimitedDiscovery", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 4) },
            { "Devices_Aep_Bluetooth_Cod_Services_Networking", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 6) },
            { "Devices_Aep_Bluetooth_Cod_Services_ObjectXfer", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 9) },
            { "Devices_Aep_Bluetooth_Cod_Services_Positioning", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 5) },
            { "Devices_Aep_Bluetooth_Cod_Services_Rendering", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 7) },
            { "Devices_Aep_Bluetooth_Cod_Services_Telephony", ("{5fbd34cd-561a-412e-ba98-478a6b0fef1d}", 11) },
            { "Devices_Aep_Bluetooth_LastSeenTime", ("{2bd67d8b-8beb-48d5-87e0-6cda3428040a}", 12) },
            { "Devices_Aep_Bluetooth_Le_AddressType", ("{995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69}", 4) },
            { "Devices_Aep_Bluetooth_Le_Appearance", ("{995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69}", 1) },
            { "Devices_Aep_Bluetooth_Le_Appearance_Category", ("{995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69}", 5) },
            { "Devices_Aep_Bluetooth_Le_Appearance_Subcategory", ("{995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69}", 6) },
            { "Devices_Aep_Bluetooth_Le_IsCallControlClient", ("{995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69}", 12) },
            { "Devices_Aep_Bluetooth_Le_IsConnectable", ("{995ef0b0-7eb3-4a8b-b9ce-068bb3f4af69}", 8) },
            { "Devices_Aep_CanPair", ("{e7c3fb29-caa7-4f47-8c8b-be59b330d4c5}", 3) },
            { "Devices_Aep_Category", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 17) },
            { "Devices_Aep_ContainerId", ("{e7c3fb29-caa7-4f47-8c8b-be59b330d4c5}", 2) },
            { "Devices_Aep_DeviceAddress", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 12) },
            { "Devices_Aep_IsConnected", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 7) },
            { "Devices_Aep_IsPaired", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 16) },
            { "Devices_Aep_IsPresent", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 9) },
            { "Devices_Aep_Manufacturer", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 5) },
            { "Devices_Aep_ModelId", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 4) },
            { "Devices_Aep_ModelName", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 3) },
            { "Devices_Aep_PointOfService_ConnectionTypes", ("{d4bf61b3-442e-4ada-882d-fa7b70c832d9}", 6) },
            { "Devices_Aep_ProtocolId", ("{3b2ce006-5e61-4fde-bab8-9b8aac9b26df}", 5) },
            { "Devices_Aep_SignalStrength", ("{a35996ab-11cf-4935-8b61-a6761081ecdf}", 6) },
            { "Devices_AepContainer_CanPair", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 3) },
            { "Devices_AepContainer_Categories", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 9) },
            { "Devices_AepContainer_Children", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 2) },
            { "Devices_AepContainer_ContainerId", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 12) },
            { "Devices_AepContainer_DialProtocol_InstalledApplications", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 6) },
            { "Devices_AepContainer_IsPaired", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 4) },
            { "Devices_AepContainer_IsPresent", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 11) },
            { "Devices_AepContainer_Manufacturer", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 6) },
            { "Devices_AepContainer_ModelIds", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 8) },
            { "Devices_AepContainer_ModelName", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 7) },
            { "Devices_AepContainer_ProtocolIds", ("{0bba1ede-7566-4f47-90ec-25fc567ced2a}", 13) },
            { "Devices_AepContainer_SupportedUriSchemes", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 5) },
            { "Devices_AepContainer_SupportsAudio", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 2) },
            { "Devices_AepContainer_SupportsCapturing", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 11) },
            { "Devices_AepContainer_SupportsImages", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 4) },
            { "Devices_AepContainer_SupportsInformation", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 14) },
            { "Devices_AepContainer_SupportsLimitedDiscovery", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 7) },
            { "Devices_AepContainer_SupportsNetworking", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 9) },
            { "Devices_AepContainer_SupportsObjectTransfer", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 12) },
            { "Devices_AepContainer_SupportsPositioning", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 8) },
            { "Devices_AepContainer_SupportsRendering", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 10) },
            { "Devices_AepContainer_SupportsTelephony", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 13) },
            { "Devices_AepContainer_SupportsVideo", ("{6af55d45-38db-4495-acb0-d4728a3b8314}", 3) },
            { "Devices_AepService_AepId", ("{c9c141a9-1b4c-4f17-a9d1-f298538cadb8}", 6) },
            { "Devices_AepService_Bluetooth_CacheMode", ("{9744311e-7951-4b2e-b6f0-ecb293cac119}", 5) },
            { "Devices_AepService_Bluetooth_ServiceGuid", ("{a399aac7-c265-474e-b073-ffce57721716}", 2) },
            { "Devices_AepService_Bluetooth_TargetDevice", ("{9744311e-7951-4b2e-b6f0-ecb293cac119}", 6) },
            { "Devices_AepService_ContainerId", ("{71724756-3e74-4432-9b59-e7b2f668a593}", 4) },
            { "Devices_AepService_FriendlyName", ("{71724756-3e74-4432-9b59-e7b2f668a593}", 2) },
            { "Devices_AepService_IoT_ServiceInterfaces", ("{79d94e82-4d79-45aa-821a-74858b4e4ca6}", 2) },
            { "Devices_AepService_ParentAepIsPaired", ("{c9c141a9-1b4c-4f17-a9d1-f298538cadb8}", 7) },
            { "Devices_AepService_ProtocolId", ("{c9c141a9-1b4c-4f17-a9d1-f298538cadb8}", 5) },
            { "Devices_AepService_ServiceClassId", ("{71724756-3e74-4432-9b59-e7b2f668a593}", 3) },
            { "Devices_AepService_ServiceId", ("{c9c141a9-1b4c-4f17-a9d1-f298538cadb8}", 2) },
            { "Devices_AppPackageFamilyName", ("{51236583-0c4a-4fe8-b81f-166aec13f510}", 100) },
            { "Devices_AudioDevice_Microphone_EqCoefficientsDb", ("{8943b373-388c-4395-b557-bc6dbaffafdb}", 7) },
            { "Devices_AudioDevice_Microphone_IsFarField", ("{8943b373-388c-4395-b557-bc6dbaffafdb}", 6) },
            { "Devices_AudioDevice_Microphone_SensitivityInDbfs", ("{8943b373-388c-4395-b557-bc6dbaffafdb}", 3) },
            { "Devices_AudioDevice_Microphone_SensitivityInDbfs2", ("{8943b373-388c-4395-b557-bc6dbaffafdb}", 5) },
            { "Devices_AudioDevice_Microphone_SignalToNoiseRatioInDb", ("{8943b373-388c-4395-b557-bc6dbaffafdb}", 4) },
            { "Devices_AudioDevice_RawProcessingSupported", ("{8943b373-388c-4395-b557-bc6dbaffafdb}", 2) },
            { "Devices_AudioDevice_SpeechProcessingSupported", ("{fb1de864-e06d-47f4-82a6-8a0aef44493c}", 2) },
            { "Devices_BatteryLife", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 10) },
            { "Devices_BatteryPlusCharging", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 22) },
            { "Devices_BatteryPlusChargingText", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 23) },
            { "Devices_Category", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 91) },
            { "Devices_CategoryGroup", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 94) },
            { "Devices_CategoryIds", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 90) },
            { "Devices_CategoryPlural", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 92) },
            { "Devices_ChallengeAep", ("{0774315e-b714-48ec-8de8-8125c077ac11}", 2) },
            { "Devices_ChargingState", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 11) },
            { "Devices_Children", ("{4340a6c5-93fa-4706-972c-7b648008a5a7}", 9) },
            { "Devices_ClassGuid", ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 10) },
            { "Devices_CompatibleIds", ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 4) },
            { "Devices_Connected", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 55) },
            { "Devices_ContainerId", ("{8c7ed206-3f8a-4827-b3ab-ae9e1faefc6c}", 2) },
            { "Devices_DefaultTooltip", ("{880f70a2-6082-47ac-8aab-a739d1a300c3}", 153) },
            { "Devices_DeviceCapabilities", ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 17) },
            { "Devices_DeviceCharacteristics", ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 29) },
            { "Devices_DeviceDescription1", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 81) },
            { "Devices_DeviceDescription2", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 82) },
            { "Devices_DeviceHasProblem", ("{540b947e-8b40-45bc-a8a2-6a0b894cbda2}", 6) },
            { "Devices_DeviceInstanceId", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 256) },
            { "Devices_DeviceManufacturer", ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 13) },
            { "Devices_DevObjectType", ("{13673f42-a3d6-49f6-b4da-ae46e0c5237c}", 2) },
            { "Devices_DialProtocol_InstalledApplications", ("{6845cc72-1b71-48c3-af86-b09171a19b14}", 3) },
            { "Devices_DiscoveryMethod", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 52) },
            { "Devices_Dnssd_Domain", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 3) },
            { "Devices_Dnssd_FullName", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 5) },
            { "Devices_Dnssd_HostName", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 7) },
            { "Devices_Dnssd_InstanceName", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 4) },
            { "Devices_Dnssd_NetworkAdapterId", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 11) },
            { "Devices_Dnssd_PortNumber", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 12) },
            { "Devices_Dnssd_Priority", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 9) },
            { "Devices_Dnssd_ServiceName", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 2) },
            { "Devices_Dnssd_TextAttributes", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 6) },
            { "Devices_Dnssd_Ttl", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 10) },
            { "Devices_Dnssd_Weight", ("{bf79c0ab-bb74-4cee-b070-470b5ae202ea}", 8) },
            { "Devices_FriendlyName", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 12288) },
            { "Devices_FunctionPaths", ("{d08dd4c0-3a9e-462e-8290-7b636b2576b9}", 3) },
            { "Devices_GlyphIcon", ("{51236583-0c4a-4fe8-b81f-166aec13f510}", 123) },
            { "Devices_HardwareIds", ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 3) },
            { "Devices_Icon", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 57) },
            { "Devices_InLocalMachineContainer", ("{8c7ed206-3f8a-4827-b3ab-ae9e1faefc6c}", 4) },
            { "Devices_InterfaceClassGuid", ("{026e516e-b814-414b-83cd-856d6fef4822}", 4) },
            { "Devices_InterfaceEnabled", ("{026e516e-b814-414b-83cd-856d6fef4822}", 3) },
            { "Devices_InterfacePaths", ("{d08dd4c0-3a9e-462e-8290-7b636b2576b9}", 2) },
            { "Devices_IpAddress", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 12297) },
            { "Devices_IsDefault", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 86) },
            { "Devices_IsNetworkConnected", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 85) },
            { "Devices_IsShared", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 84) },
            { "Devices_IsSoftwareInstalling", ("{83da6326-97a6-4088-9453-a1923f573b29}", 9) },
            { "Devices_LaunchDeviceStageFromExplorer", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 77) },
            { "Devices_LocalMachine", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 70) },
            { "Devices_LocationPaths", ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 37) },
            { "Devices_Manufacturer", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 8192) },
            { "Devices_MetadataPath", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 71) },
            { "Devices_MicrophoneArray_Geometry", ("{a1829ea2-27eb-459e-935d-b2fad7b07762}", 2) },
            { "Devices_MissedCalls", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 5) },
            { "Devices_ModelId", ("{80d81ea6-7473-4b0c-8216-efc11a2c4c8b}", 2) },
            { "Devices_ModelName", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 8194) },
            { "Devices_ModelNumber", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 8195) },
            { "Devices_NetworkedTooltip", ("{880f70a2-6082-47ac-8aab-a739d1a300c3}", 152) },
            { "Devices_NetworkName", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 7) },
            { "Devices_NetworkType", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 8) },
            { "Devices_NewPictures", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 4) },
            { "Devices_Notification", ("{06704b0c-e830-4c81-9178-91e4e95a80a0}", 3) },
            { "Devices_Notifications_LowBattery", ("{c4c07f2b-8524-4e66-ae3a-a6235f103beb}", 2) },
            { "Devices_Notifications_MissedCall", ("{6614ef48-4efe-4424-9eda-c79f404edf3e}", 2) },
            { "Devices_Notifications_NewMessage", ("{2be9260a-2012-4742-a555-f41b638b7dcb}", 2) },
            { "Devices_Notifications_NewVoicemail", ("{59569556-0a08-4212-95b9-fae2ad6413db}", 2) },
            { "Devices_Notifications_StorageFull", ("{a0e00ee1-f0c7-4d41-b8e7-26a7bd8d38b0}", 2) },
            { "Devices_Notifications_StorageFullLinkText", ("{a0e00ee1-f0c7-4d41-b8e7-26a7bd8d38b0}", 3) },
            { "Devices_NotificationStore", ("{06704b0c-e830-4c81-9178-91e4e95a80a0}", 2) },
            { "Devices_NotWorkingProperly", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 83) },
            { "Devices_Paired", ("{78c34fc8-104a-4aca-9ea4-524d52996e57}", 56) },
            { "Devices_Panel_PanelGroup", ("{8dbc9c86-97a9-4bff-9bc6-bfe95d3e6dad}", 3) },
            { "Devices_Panel_PanelId", ("{8dbc9c86-97a9-4bff-9bc6-bfe95d3e6dad}", 2) },
            { "Devices_Parent", ("{4340a6c5-93fa-4706-972c-7b648008a5a7}", 8) },
            { "Devices_PhoneLineTransportDevice_Connected", ("{aecf2fe8-1d00-4fee-8a6d-a70d719b772b}", 2) },
            { "Devices_PhysicalDeviceLocation", ("{540b947e-8b40-45bc-a8a2-6a0b894cbda2}", 9) },
            { "Devices_PlaybackPositionPercent", ("{3633de59-6825-4381-a49b-9f6ba13a1471}", 5) },
            { "Devices_PlaybackState", ("{3633de59-6825-4381-a49b-9f6ba13a1471}", 2) },
            { "Devices_PlaybackTitle", ("{3633de59-6825-4381-a49b-9f6ba13a1471}", 3) },
            { "Devices_Present", ("{540b947e-8b40-45bc-a8a2-6a0b894cbda2}", 5) },
            { "Devices_PresentationUrl", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 8198) },
            { "Devices_PrimaryCategory", ("{d08dd4c0-3a9e-462e-8290-7b636b2576b9}", 10) },
            { "Devices_RemainingDuration", ("{3633de59-6825-4381-a49b-9f6ba13a1471}", 4) },
            { "Devices_RestrictedInterface", ("{026e516e-b814-414b-83cd-856d6fef4822}", 6) },
            { "Devices_Roaming", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 9) },
            { "Devices_SafeRemovalRequired", ("{afd97640-86a3-4210-b67c-289c41aabe55}", 2) },
            { "Devices_SchematicName", ("{026e516e-b814-414b-83cd-856d6fef4822}", 9) },
            { "Devices_ServiceAddress", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 16384) },
            { "Devices_ServiceId", ("{656a3bb3-ecc0-43fd-8477-4ae0404a96cd}", 16385) },
            { "Devices_SharedTooltip", ("{880f70a2-6082-47ac-8aab-a739d1a300c3}", 151) },
            { "Devices_SignalStrength", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 2) },
            { "Devices_SmartCards_ReaderKind", ("{d6b5b883-18bd-4b4d-b2ec-9e38affeda82}", 2) },
            { "Devices_Status", ("{d08dd4c0-3a9e-462e-8290-7b636b2576b9}", 259) },
            { "Devices_Status1", ("{d08dd4c0-3a9e-462e-8290-7b636b2576b9}", 257) },
            { "Devices_Status2", ("{d08dd4c0-3a9e-462e-8290-7b636b2576b9}", 258) },
            { "Devices_StorageCapacity", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 12) },
            { "Devices_StorageFreeSpace", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 13) },
            { "Devices_StorageFreeSpacePercent", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 14) },
            { "Devices_TextMessages", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 3) },
            { "Devices_Voicemail", ("{49cd1f76-5626-4b17-a4e8-18b4aa1a2213}", 6) },
            { "Devices_WiaDeviceType", ("{6bdd1fc6-810f-11d0-bec7-08002be2092f}", 2) },
            { "Devices_WiFi_InterfaceGuid", ("{ef1167eb-cbfc-4341-a568-a7c91a68982c}", 2) },
            { "Devices_WiFiDirect_DeviceAddress", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 13) },
            { "Devices_WiFiDirect_GroupId", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 4) },
            { "Devices_WiFiDirect_InformationElements", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 12) },
            { "Devices_WiFiDirect_InterfaceAddress", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 2) },
            { "Devices_WiFiDirect_InterfaceGuid", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 3) },
            { "Devices_WiFiDirect_IsConnected", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 5) },
            { "Devices_WiFiDirect_IsLegacyDevice", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 7) },
            { "Devices_WiFiDirect_IsMiracastLcpSupported", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 9) },
            { "Devices_WiFiDirect_IsVisible", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 6) },
            { "Devices_WiFiDirect_MiracastVersion", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 8) },
            { "Devices_WiFiDirect_Services", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 10) },
            { "Devices_WiFiDirect_SupportedChannelList", ("{1506935d-e3e7-450f-8637-82233ebe5f6e}", 11) },
            { "Devices_WiFiDirectServices_AdvertisementId", ("{31b37743-7c5e-4005-93e6-e953f92b82e9}", 5) },
            { "Devices_WiFiDirectServices_RequestServiceInformation", ("{31b37743-7c5e-4005-93e6-e953f92b82e9}", 7) },
            { "Devices_WiFiDirectServices_ServiceAddress", ("{31b37743-7c5e-4005-93e6-e953f92b82e9}", 2) },
            { "Devices_WiFiDirectServices_ServiceConfigMethods", ("{31b37743-7c5e-4005-93e6-e953f92b82e9}", 6) },
            { "Devices_WiFiDirectServices_ServiceInformation", ("{31b37743-7c5e-4005-93e6-e953f92b82e9}", 4) },
            { "Devices_WiFiDirectServices_ServiceName", ("{31b37743-7c5e-4005-93e6-e953f92b82e9}", 3) },
            { "Devices_WinPhone8CameraFlags", ("{b7b4d61c-5a64-4187-a52e-b1539f359099}", 2) },
            { "Devices_Wwan_InterfaceGuid", ("{ff1167eb-cbfc-4341-a568-a7c91a68982c}", 2) },
            { "Storage_Portable", ("{4d1ebee8-0803-4774-9842-b77db50265e9}", 2) },
            { "Storage_RemovableMedia", ("{4d1ebee8-0803-4774-9842-b77db50265e9}", 3) },
            { "Storage_SystemCritical", ("{4d1ebee8-0803-4774-9842-b77db50265e9}", 4) },
            { "Document_ByteCount", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 4) },
            { "Document_CharacterCount", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 16) },
            { "Document_ClientID", ("{276d7bb0-5b34-4fb0-aa4b-158ed12a1809}", 100) },
            { "Document_Contributor", ("{f334115e-da1b-4509-9b3d-119504dc7abb}", 100) },
            { "Document_DateCreated", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 12) },
            { "Document_DatePrinted", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 11) },
            { "Document_DateSaved", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 13) },
            { "Document_Division", ("{1e005ee6-bf27-428b-b01c-79676acd2870}", 100) },
            { "Document_DocumentID", ("{e08805c8-e395-40df-80d2-54f0d6c43154}", 100) },
            { "Document_HiddenSlideCount", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 9) },
            { "Document_LastAuthor", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 8) },
            { "Document_LineCount", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 5) },
            { "Document_Manager", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 14) },
            { "Document_MultimediaClipCount", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 10) },
            { "Document_NoteCount", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 8) },
            { "Document_PageCount", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 14) },
            { "Document_ParagraphCount", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 6) },
            { "Document_PresentationFormat", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 3) },
            { "Document_RevisionNumber", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 9) },
            { "Document_Security", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 19) },
            { "Document_SlideCount", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 7) },
            { "Document_Template", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 7) },
            { "Document_TotalEditingTime", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 10) },
            { "Document_Version", ("{d5cdd502-2e9c-101b-9397-08002b2cf9ae}", 29) },
            { "Document_WordCount", ("{f29f85e0-4ff9-1068-ab91-08002b27b3d9}", 15) },
            { "DRM_DatePlayExpires", ("{aeac19e4-89ae-4508-b9b7-bb867abee2ed}", 6) },
            { "DRM_DatePlayStarts", ("{aeac19e4-89ae-4508-b9b7-bb867abee2ed}", 5) },
            { "DRM_Description", ("{aeac19e4-89ae-4508-b9b7-bb867abee2ed}", 3) },
            { "DRM_IsDisabled", ("{aeac19e4-89ae-4508-b9b7-bb867abee2ed}", 7) },
            { "DRM_IsProtected", ("{aeac19e4-89ae-4508-b9b7-bb867abee2ed}", 2) },
            { "DRM_PlayCount", ("{aeac19e4-89ae-4508-b9b7-bb867abee2ed}", 4) },
            { "GPS_Altitude", ("{827edb4f-5b73-44a7-891d-fdffabea35ca}", 100) },
            { "GPS_AltitudeDenominator", ("{78342dcb-e358-4145-ae9a-6bfe4e0f9f51}", 100) },
            { "GPS_AltitudeNumerator", ("{2dad1eb7-816d-40d3-9ec3-c9773be2aade}", 100) },
            { "GPS_AltitudeRef", ("{46ac629d-75ea-4515-867f-6dc4321c5844}", 100) },
            { "GPS_AreaInformation", ("{972e333e-ac7e-49f1-8adf-a70d07a9bcab}", 100) },
            { "GPS_Date", ("{3602c812-0f3b-45f0-85ad-603468d69423}", 100) },
            { "GPS_DestBearing", ("{c66d4b3c-e888-47cc-b99f-9dca3ee34dea}", 100) },
            { "GPS_DestBearingDenominator", ("{7abcf4f8-7c3f-4988-ac91-8d2c2e97eca5}", 100) },
            { "GPS_DestBearingNumerator", ("{ba3b1da9-86ee-4b5d-a2a4-a271a429f0cf}", 100) },
            { "GPS_DestBearingRef", ("{9ab84393-2a0f-4b75-bb22-7279786977cb}", 100) },
            { "GPS_DestDistance", ("{a93eae04-6804-4f24-ac81-09b266452118}", 100) },
            { "GPS_DestDistanceDenominator", ("{9bc2c99b-ac71-4127-9d1c-2596d0d7dcb7}", 100) },
            { "GPS_DestDistanceNumerator", ("{2bda47da-08c6-4fe1-80bc-a72fc517c5d0}", 100) },
            { "GPS_DestDistanceRef", ("{ed4df2d3-8695-450b-856f-f5c1c53acb66}", 100) },
            { "GPS_DestLatitude", ("{9d1d7cc5-5c39-451c-86b3-928e2d18cc47}", 100) },
            { "GPS_DestLatitudeDenominator", ("{3a372292-7fca-49a7-99d5-e47bb2d4e7ab}", 100) },
            { "GPS_DestLatitudeNumerator", ("{ecf4b6f6-d5a6-433c-bb92-4076650fc890}", 100) },
            { "GPS_DestLatitudeRef", ("{cea820b9-ce61-4885-a128-005d9087c192}", 100) },
            { "GPS_DestLongitude", ("{47a96261-cb4c-4807-8ad3-40b9d9dbc6bc}", 100) },
            { "GPS_DestLongitudeDenominator", ("{425d69e5-48ad-4900-8d80-6eb6b8d0ac86}", 100) },
            { "GPS_DestLongitudeNumerator", ("{a3250282-fb6d-48d5-9a89-dbcace75cccf}", 100) },
            { "GPS_DestLongitudeRef", ("{182c1ea6-7c1c-4083-ab4b-ac6c9f4ed128}", 100) },
            { "GPS_Differential", ("{aaf4ee25-bd3b-4dd7-bfc4-47f77bb00f6d}", 100) },
            { "GPS_DOP", ("{0cf8fb02-1837-42f1-a697-a7017aa289b9}", 100) },
            { "GPS_DOPDenominator", ("{a0be94c5-50ba-487b-bd35-0654be8881ed}", 100) },
            { "GPS_DOPNumerator", ("{47166b16-364f-4aa0-9f31-e2ab3df449c3}", 100) },
            { "GPS_ImgDirection", ("{16473c91-d017-4ed9-ba4d-b6baa55dbcf8}", 100) },
            { "GPS_ImgDirectionDenominator", ("{10b24595-41a2-4e20-93c2-5761c1395f32}", 100) },
            { "GPS_ImgDirectionNumerator", ("{dc5877c7-225f-45f7-bac7-e81334b6130a}", 100) },
            { "GPS_ImgDirectionRef", ("{a4aaa5b7-1ad0-445f-811a-0f8f6e67f6b5}", 100) },
            { "GPS_Latitude", ("{8727cfff-4868-4ec6-ad5b-81b98521d1ab}", 100) },
            { "GPS_LatitudeDecimal", ("{0f55cde2-4f49-450d-92c1-dcd16301b1b7}", 100) },
            { "GPS_LatitudeDenominator", ("{16e634ee-2bff-497b-bd8a-4341ad39eeb9}", 100) },
            { "GPS_LatitudeNumerator", ("{7ddaaad1-ccc8-41ae-b750-b2cb8031aea2}", 100) },
            { "GPS_LatitudeRef", ("{029c0252-5b86-46c7-aca0-2769ffc8e3d4}", 100) },
            { "GPS_Longitude", ("{c4c4dbb2-b593-466b-bbda-d03d27d5e43a}", 100) },
            { "GPS_LongitudeDecimal", ("{4679c1b5-844d-4590-baf5-f322231f1b81}", 100) },
            { "GPS_LongitudeDenominator", ("{be6e176c-4534-4d2c-ace5-31dedac1606b}", 100) },
            { "GPS_LongitudeNumerator", ("{02b0f689-a914-4e45-821d-1dda452ed2c4}", 100) },
            { "GPS_LongitudeRef", ("{33dcf22b-28d5-464c-8035-1ee9efd25278}", 100) },
            { "GPS_MapDatum", ("{2ca2dae6-eddc-407d-bef1-773942abfa95}", 100) },
            { "GPS_MeasureMode", ("{a015ed5d-aaea-4d58-8a86-3c586920ea0b}", 100) },
            { "GPS_ProcessingMethod", ("{59d49e61-840f-4aa9-a939-e2099b7f6399}", 100) },
            { "GPS_Satellites", ("{467ee575-1f25-4557-ad4e-b8b58b0d9c15}", 100) },
            { "GPS_Speed", ("{da5d0862-6e76-4e1b-babd-70021bd25494}", 100) },
            { "GPS_SpeedDenominator", ("{7d122d5a-ae5e-4335-8841-d71e7ce72f53}", 100) },
            { "GPS_SpeedNumerator", ("{acc9ce3d-c213-4942-8b48-6d0820f21c6d}", 100) },
            { "GPS_SpeedRef", ("{ecf7f4c9-544f-4d6d-9d98-8ad79adaf453}", 100) },
            { "GPS_Status", ("{125491f4-818f-46b2-91b5-d537753617b2}", 100) },
            { "GPS_Track", ("{76c09943-7c33-49e3-9e7e-cdba872cfada}", 100) },
            { "GPS_TrackDenominator", ("{c8d1920c-01f6-40c0-ac86-2f3a4ad00770}", 100) },
            { "GPS_TrackNumerator", ("{702926f4-44a6-43e1-ae71-45627116893b}", 100) },
            { "GPS_TrackRef", ("{35dbe6fe-44c3-4400-aaae-d2c799c407e8}", 100) },
            { "GPS_VersionID", ("{22704da4-c6b2-4a99-8e56-f16df8c92599}", 100) },
            { "History_VisitCount", ("{5cbf2787-48cf-4208-b90e-ee5e5d420294}", 7) },
            { "Image_BitDepth", ("{6444048f-4c8b-11d1-8b70-080036b11a03}", 7) },
            { "Image_ColorSpace", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 40961) },
            { "Image_CompressedBitsPerPixel", ("{364b6fa9-37ab-482a-be2b-ae02f60d4318}", 100) },
            { "Image_CompressedBitsPerPixelDenominator", ("{1f8844e1-24ad-4508-9dfd-5326a415ce02}", 100) },
            { "Image_CompressedBitsPerPixelNumerator", ("{d21a7148-d32c-4624-8900-277210f79c0f}", 100) },
            { "Image_Compression", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 259) },
            { "Image_CompressionText", ("{3f08e66f-2f44-4bb9-a682-ac35d2562322}", 100) },
            { "Image_Dimensions", ("{6444048f-4c8b-11d1-8b70-080036b11a03}", 13) },
            { "Image_HorizontalResolution", ("{6444048f-4c8b-11d1-8b70-080036b11a03}", 5) },
            { "Image_HorizontalSize", ("{6444048f-4c8b-11d1-8b70-080036b11a03}", 3) },
            { "Image_ImageID", ("{10dabe05-32aa-4c29-bf1a-63e2d220587f}", 100) },
            { "Image_ResolutionUnit", ("{19b51fa6-1f92-4a5c-ab48-7df0abd67444}", 100) },
            { "Image_VerticalResolution", ("{6444048f-4c8b-11d1-8b70-080036b11a03}", 6) },
            { "Image_VerticalSize", ("{6444048f-4c8b-11d1-8b70-080036b11a03}", 4) },
            { "Journal_Contacts", ("{dea7c82c-1d89-4a66-9427-a4e3debabcb1}", 100) },
            { "Journal_EntryType", ("{95beb1fc-326d-4644-b396-cd3ed90e6ddf}", 100) },
            { "LayoutPattern_ContentViewModeForBrowse", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 500) },
            { "LayoutPattern_ContentViewModeForSearch", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 501) },
            { "History_SelectionCount", ("{1ce0d6bc-536c-4600-b0dd-7e0c66b350d5}", 8) },
            { "History_TargetUrlHostName", ("{1ce0d6bc-536c-4600-b0dd-7e0c66b350d5}", 9) },
            { "Link_Arguments", ("{436f2667-14e2-4feb-b30a-146c53b5b674}", 100) },
            { "Link_Comment", ("{b9b4b3fc-2b51-4a42-b5d8-324146afcf25}", 5) },
            { "Link_DateVisited", ("{5cbf2787-48cf-4208-b90e-ee5e5d420294}", 23) },
            { "Link_Description", ("{5cbf2787-48cf-4208-b90e-ee5e5d420294}", 21) },
            { "Link_FeedItemLocalId", ("{8a2f99f9-3c37-465d-a8d7-69777a246d0c}", 2) },
            { "Link_Status", ("{b9b4b3fc-2b51-4a42-b5d8-324146afcf25}", 3) },
            { "Link_TargetExtension", ("{7a7d76f4-b630-4bd7-95ff-37cc51a975c9}", 2) },
            { "Link_TargetParsingPath", ("{b9b4b3fc-2b51-4a42-b5d8-324146afcf25}", 2) },
            { "Link_TargetSFGAOFlags", ("{b9b4b3fc-2b51-4a42-b5d8-324146afcf25}", 8) },
            { "Link_TargetUrlHostName", ("{8a2f99f9-3c37-465d-a8d7-69777a246d0c}", 5) },
            { "Link_TargetUrlPath", ("{8a2f99f9-3c37-465d-a8d7-69777a246d0c}", 6) },
            { "Media_AuthorUrl", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 32) },
            { "Media_AverageLevel", ("{09edd5b6-b301-43c5-9990-d00302effd46}", 100) },
            { "Media_ClassPrimaryID", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 13) },
            { "Media_ClassSecondaryID", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 14) },
            { "Media_CollectionGroupID", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 24) },
            { "Media_CollectionID", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 25) },
            { "Media_ContentDistributor", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 18) },
            { "Media_ContentID", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 26) },
            { "Media_CreatorApplication", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 27) },
            { "Media_CreatorApplicationVersion", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 28) },
            { "Media_DateEncoded", ("{2e4b640d-5019-46d8-8881-55414cc5caa0}", 100) },
            { "Media_DateReleased", ("{de41cc29-6971-4290-b472-f59f2e2f31e2}", 100) },
            { "Media_DlnaProfileID", ("{cfa31b45-525d-4998-bb44-3f7d81542fa4}", 100) },
            { "Media_Duration", ("{64440490-4c8b-11d1-8b70-080036b11a03}", 3) },
            { "Media_DVDID", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 15) },
            { "Media_EncodedBy", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 36) },
            { "Media_EncodingSettings", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 37) },
            { "Media_EpisodeNumber", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 100) },
            { "Media_FrameCount", ("{6444048f-4c8b-11d1-8b70-080036b11a03}", 12) },
            { "Media_MCDI", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 16) },
            { "Media_MetadataContentProvider", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 17) },
            { "Media_Producer", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 22) },
            { "Media_PromotionUrl", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 33) },
            { "Media_ProtectionType", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 38) },
            { "Media_ProviderRating", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 39) },
            { "Media_ProviderStyle", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 40) },
            { "Media_Publisher", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 30) },
            { "Media_SeasonNumber", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 101) },
            { "Media_SeriesName", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 42) },
            { "Media_SubscriptionContentId", ("{9aebae7a-9644-487d-a92c-657585ed751a}", 100) },
            { "Media_SubTitle", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 38) },
            { "Media_ThumbnailLargePath", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 47) },
            { "Media_ThumbnailLargeUri", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 48) },
            { "Media_ThumbnailSmallPath", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 49) },
            { "Media_ThumbnailSmallUri", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 50) },
            { "Media_UniqueFileIdentifier", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 35) },
            { "Media_UserNoAutoInfo", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 41) },
            { "Media_UserWebUrl", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 34) },
            { "Media_Writer", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 23) },
            { "Media_Year", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 5) },
            { "Message_AttachmentContents", ("{3143bf7c-80a8-4854-8880-e2e40189bdd0}", 100) },
            { "Message_AttachmentNames", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 21) },
            { "Message_BccAddress", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 2) },
            { "Message_BccName", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 3) },
            { "Message_CcAddress", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 4) },
            { "Message_CcName", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 5) },
            { "Message_ConversationID", ("{dc8f80bd-af1e-4289-85b6-3dfc1b493992}", 100) },
            { "Message_ConversationIndex", ("{dc8f80bd-af1e-4289-85b6-3dfc1b493992}", 101) },
            { "Message_DateReceived", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 20) },
            { "Message_DateSent", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 19) },
            { "Message_Flags", ("{a82d9ee7-ca67-4312-965e-226bcea85023}", 100) },
            { "Message_FromAddress", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 13) },
            { "Message_FromName", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 14) },
            { "Message_HasAttachments", ("{9c1fcf74-2d97-41ba-b4ae-cb2e3661a6e4}", 8) },
            { "Message_IsFwdOrReply", ("{9a9bc088-4f6d-469e-9919-e705412040f9}", 100) },
            { "Message_MessageClass", ("{cd9ed458-08ce-418f-a70e-f912c7bb9c5c}", 103) },
            { "Message_Participants", ("{1a9ba605-8e7c-4d11-ad7d-a50ada18ba1b}", 2) },
            { "Message_ProofInProgress", ("{9098f33c-9a7d-48a8-8de5-2e1227a64e91}", 100) },
            { "Message_SenderAddress", ("{0be1c8e7-1981-4676-ae14-fdd78f05a6e7}", 100) },
            { "Message_SenderName", ("{0da41cfa-d224-4a18-ae2f-596158db4b3a}", 100) },
            { "Message_Store", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 15) },
            { "Message_ToAddress", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 16) },
            { "Message_ToDoFlags", ("{1f856a9f-6900-4aba-9505-2d5f1b4d66cb}", 100) },
            { "Message_ToDoTitle", ("{bccc8a3c-8cef-42e5-9b1c-c69079398bc7}", 100) },
            { "Message_ToName", ("{e3e0584c-b788-4a5a-bb20-7f5a44c9acdd}", 17) },
            { "MsGraph_ActivityType", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 14) },
            { "MsGraph_CompositeId", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 2) },
            { "MsGraph_DateLastShared", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 9) },
            { "MsGraph_DriveId", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 3) },
            { "MsGraph_GraphFileType", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 16) },
            { "MsGraph_IconUrl", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 15) },
            { "MsGraph_ItemId", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 4) },
            { "MsGraph_PrimaryActivityActorDisplayName", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 13) },
            { "MsGraph_PrimaryActivityActorUpn", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 12) },
            { "MsGraph_RecommendationReason", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 8) },
            { "MsGraph_RecommendationReferenceId", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 5) },
            { "MsGraph_RecommendationResultSourceId", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 7) },
            { "MsGraph_SharedByEmail", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 11) },
            { "MsGraph_SharedByName", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 10) },
            { "MsGraph_WebAccountId", ("{4f85567e-fff0-4df5-b1d9-98b314ff0729}", 6) },
            { "Music_AlbumArtist", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 13) },
            { "Music_AlbumArtistSortOverride", ("{f1fdb4af-f78c-466c-bb05-56e92db0b8ec}", 103) },
            { "Music_AlbumID", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 100) },
            { "Music_AlbumTitle", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 4) },
            { "Music_AlbumTitleSortOverride", ("{13eb7ffc-ec89-4346-b19d-ccc6f1784223}", 101) },
            { "Music_Artist", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 2) },
            { "Music_ArtistSortOverride", ("{deeb2db5-0696-4ce0-94fe-a01f77a45fb5}", 102) },
            { "Music_BeatsPerMinute", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 35) },
            { "Music_Composer", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 19) },
            { "Music_ComposerSortOverride", ("{00bc20a3-bd48-4085-872c-a88d77f5097e}", 105) },
            { "Music_Conductor", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 36) },
            { "Music_ContentGroupDescription", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 33) },
            { "Music_DiscNumber", ("{6afe7437-9bcd-49c7-80fe-4a5c65fa5874}", 104) },
            { "Music_DisplayArtist", ("{fd122953-fa93-4ef7-92c3-04c946b2f7c8}", 100) },
            { "Music_Genre", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 11) },
            { "Music_InitialKey", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 34) },
            { "Music_IsCompilation", ("{c449d5cb-9ea4-4809-82e8-af9d59ded6d1}", 100) },
            { "Music_Lyrics", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 12) },
            { "Music_Mood", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 39) },
            { "Music_PartOfSet", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 37) },
            { "Music_Period", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 31) },
            { "Music_SynchronizedLyrics", ("{6b223b6a-162e-4aa9-b39f-05d678fc6d77}", 100) },
            { "Music_TrackNumber", ("{56a3372e-ce9c-11d2-9f0e-006097c686f6}", 7) },
            { "Note_Color", ("{4776cafa-bce4-4cb1-a23e-265e76d8eb11}", 100) },
            { "Note_ColorText", ("{46b4e8de-cdb2-440d-885c-1658eb65b914}", 100) },
            { "Photo_Aperture", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37378) },
            { "Photo_ApertureDenominator", ("{e1a9a38b-6685-46bd-875e-570dc7ad7320}", 100) },
            { "Photo_ApertureNumerator", ("{0337ecec-39fb-4581-a0bd-4c4cc51e9914}", 100) },
            { "Photo_Brightness", ("{1a701bf6-478c-4361-83ab-3701bb053c58}", 100) },
            { "Photo_BrightnessDenominator", ("{6ebe6946-2321-440a-90f0-c043efd32476}", 100) },
            { "Photo_BrightnessNumerator", ("{9e7d118f-b314-45a0-8cfb-d654b917c9e9}", 100) },
            { "Photo_CameraManufacturer", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 271) },
            { "Photo_CameraModel", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 272) },
            { "Photo_CameraSerialNumber", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 273) },
            { "Photo_Contrast", ("{2a785ba9-8d23-4ded-82e6-60a350c86a10}", 100) },
            { "Photo_ContrastText", ("{59dde9f2-5253-40ea-9a8b-479e96c6249a}", 100) },
            { "Photo_DateTaken", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 36867) },
            { "Photo_DigitalZoom", ("{f85bf840-a925-4bc2-b0c4-8e36b598679e}", 100) },
            { "Photo_DigitalZoomDenominator", ("{745baf0e-e5c1-4cfb-8a1b-d031a0a52393}", 100) },
            { "Photo_DigitalZoomNumerator", ("{16cbb924-6500-473b-a5be-f1599bcbe413}", 100) },
            { "Photo_Event", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 18248) },
            { "Photo_EXIFVersion", ("{d35f743a-eb2e-47f2-a286-844132cb1427}", 100) },
            { "Photo_ExposureBias", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37380) },
            { "Photo_ExposureBiasDenominator", ("{ab205e50-04b7-461c-a18c-2f233836e627}", 100) },
            { "Photo_ExposureBiasNumerator", ("{738bf284-1d87-420b-92cf-5834bf6ef9ed}", 100) },
            { "Photo_ExposureIndex", ("{967b5af8-995a-46ed-9e11-35b3c5b9782d}", 100) },
            { "Photo_ExposureIndexDenominator", ("{93112f89-c28b-492f-8a9d-4be2062cee8a}", 100) },
            { "Photo_ExposureIndexNumerator", ("{cdedcf30-8919-44df-8f4c-4eb2ffdb8d89}", 100) },
            { "Photo_ExposureProgram", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 34850) },
            { "Photo_ExposureProgramText", ("{fec690b7-5f30-4646-ae47-4caafba884a3}", 100) },
            { "Photo_ExposureTime", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 33434) },
            { "Photo_ExposureTimeDenominator", ("{55e98597-ad16-42e0-b624-21599a199838}", 100) },
            { "Photo_ExposureTimeNumerator", ("{257e44e2-9031-4323-ac38-85c552871b2e}", 100) },
            { "Photo_Flash", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37385) },
            { "Photo_FlashEnergy", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 41483) },
            { "Photo_FlashEnergyDenominator", ("{d7b61c70-6323-49cd-a5fc-c84277162c97}", 100) },
            { "Photo_FlashEnergyNumerator", ("{fcad3d3d-0858-400f-aaa3-2f66cce2a6bc}", 100) },
            { "Photo_FlashManufacturer", ("{aabaf6c9-e0c5-4719-8585-57b103e584fe}", 100) },
            { "Photo_FlashModel", ("{fe83bb35-4d1a-42e2-916b-06f3e1af719e}", 100) },
            { "Photo_FlashText", ("{6b8b68f6-200b-47ea-8d25-d8050f57339f}", 100) },
            { "Photo_FNumber", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 33437) },
            { "Photo_FNumberDenominator", ("{e92a2496-223b-4463-a4e3-30eabba79d80}", 100) },
            { "Photo_FNumberNumerator", ("{1b97738a-fdfc-462f-9d93-1957e08be90c}", 100) },
            { "Photo_FocalLength", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37386) },
            { "Photo_FocalLengthDenominator", ("{305bc615-dca1-44a5-9fd4-10c0ba79412e}", 100) },
            { "Photo_FocalLengthInFilm", ("{a0e74609-b84d-4f49-b860-462bd9971f98}", 100) },
            { "Photo_FocalLengthNumerator", ("{776b6b3b-1e3d-4b0c-9a0e-8fbaf2a8492a}", 100) },
            { "Photo_FocalPlaneXResolution", ("{cfc08d97-c6f7-4484-89dd-ebef4356fe76}", 100) },
            { "Photo_FocalPlaneXResolutionDenominator", ("{0933f3f5-4786-4f46-a8e8-d64dd37fa521}", 100) },
            { "Photo_FocalPlaneXResolutionNumerator", ("{dccb10af-b4e2-4b88-95f9-031b4d5ab490}", 100) },
            { "Photo_FocalPlaneYResolution", ("{4fffe4d0-914f-4ac4-8d6f-c9c61de169b1}", 100) },
            { "Photo_FocalPlaneYResolutionDenominator", ("{1d6179a6-a876-4031-b013-3347b2b64dc8}", 100) },
            { "Photo_FocalPlaneYResolutionNumerator", ("{a2e541c5-4440-4ba8-867e-75cfc06828cd}", 100) },
            { "Photo_GainControl", ("{fa304789-00c7-4d80-904a-1e4dcc7265aa}", 100) },
            { "Photo_GainControlDenominator", ("{42864dfd-9da4-4f77-bded-4aad7b256735}", 100) },
            { "Photo_GainControlNumerator", ("{8e8ecf7c-b7b8-4eb8-a63f-0ee715c96f9e}", 100) },
            { "Photo_GainControlText", ("{c06238b2-0bf9-4279-a723-25856715cb9d}", 100) },
            { "Photo_ISOSpeed", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 34855) },
            { "Photo_LensManufacturer", ("{e6ddcaf7-29c5-4f0a-9a68-d19412ec7090}", 100) },
            { "Photo_LensModel", ("{e1277516-2b5f-4869-89b1-2e585bd38b7a}", 100) },
            { "Photo_LightSource", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37384) },
            { "Photo_MakerNote", ("{fa303353-b659-4052-85e9-bcac79549b84}", 100) },
            { "Photo_MakerNoteOffset", ("{813f4124-34e6-4d17-ab3e-6b1f3c2247a1}", 100) },
            { "Photo_MaxAperture", ("{08f6d7c2-e3f2-44fc-af1e-5aa5c81a2d3e}", 100) },
            { "Photo_MaxApertureDenominator", ("{c77724d4-601f-46c5-9b89-c53f93bceb77}", 100) },
            { "Photo_MaxApertureNumerator", ("{c107e191-a459-44c5-9ae6-b952ad4b906d}", 100) },
            { "Photo_MeteringMode", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37383) },
            { "Photo_MeteringModeText", ("{f628fd8c-7ba8-465a-a65b-c5aa79263a9e}", 100) },
            { "Photo_Orientation", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 274) },
            { "Photo_OrientationText", ("{a9ea193c-c511-498a-a06b-58e2776dcc28}", 100) },
            { "Photo_PeopleNames", ("{e8309b6e-084c-49b4-b1fc-90a80331b638}", 100) },
            { "Photo_PhotometricInterpretation", ("{341796f1-1df9-4b1c-a564-91bdefa43877}", 100) },
            { "Photo_PhotometricInterpretationText", ("{821437d6-9eab-4765-a589-3b1cbbd22a61}", 100) },
            { "Photo_ProgramMode", ("{6d217f6d-3f6a-4825-b470-5f03ca2fbe9b}", 100) },
            { "Photo_ProgramModeText", ("{7fe3aa27-2648-42f3-89b0-454e5cb150c3}", 100) },
            { "Photo_RelatedSoundFile", ("{318a6b45-087f-4dc2-b8cc-05359551fc9e}", 100) },
            { "Photo_Saturation", ("{49237325-a95a-4f67-b211-816b2d45d2e0}", 100) },
            { "Photo_SaturationText", ("{61478c08-b600-4a84-bbe4-e99c45f0a072}", 100) },
            { "Photo_Sharpness", ("{fc6976db-8349-4970-ae97-b3c5316a08f0}", 100) },
            { "Photo_SharpnessText", ("{51ec3f47-dd50-421d-8769-334f50424b1e}", 100) },
            { "Photo_ShutterSpeed", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37377) },
            { "Photo_ShutterSpeedDenominator", ("{e13d8975-81c7-4948-ae3f-37cae11e8ff7}", 100) },
            { "Photo_ShutterSpeedNumerator", ("{16ea4042-d6f4-4bca-8349-7c78d30fb333}", 100) },
            { "Photo_SubjectDistance", ("{14b81da1-0135-4d31-96d9-6cbfc9671a99}", 37382) },
            { "Photo_SubjectDistanceDenominator", ("{0c840a88-b043-466d-9766-d4b26da3fa77}", 100) },
            { "Photo_SubjectDistanceNumerator", ("{8af4961c-f526-43e5-aa81-db768219178d}", 100) },
            { "Photo_TagViewAggregate", ("{b812f15d-c2d8-4bbf-bacd-79744346113f}", 100) },
            { "Photo_TranscodedForSync", ("{9a8ebb75-6458-4e82-bacb-35c0095b03bb}", 100) },
            { "Photo_WhiteBalance", ("{ee3d3d8a-5381-4cfa-b13b-aaf66b5f4ec9}", 100) },
            { "Photo_WhiteBalanceText", ("{6336b95e-c7a7-426d-86fd-7ae3d39c84b4}", 100) },
            { "PropGroup_Advanced", ("{900a403b-097b-4b95-8ae2-071fdaeeb118}", 100) },
            { "PropGroup_Audio", ("{2804d469-788f-48aa-8570-71b9c187e138}", 100) },
            { "PropGroup_Calendar", ("{9973d2b5-bfd8-438a-ba94-5349b293181a}", 100) },
            { "PropGroup_Camera", ("{de00de32-547e-4981-ad4b-542f2e9007d8}", 100) },
            { "PropGroup_Contact", ("{df975fd3-250a-4004-858f-34e29a3e37aa}", 100) },
            { "PropGroup_Content", ("{d0dab0ba-368a-4050-a882-6c010fd19a4f}", 100) },
            { "PropGroup_Description", ("{8969b275-9475-4e00-a887-ff93b8b41e44}", 100) },
            { "PropGroup_FileSystem", ("{e3a7d2c1-80fc-4b40-8f34-30ea111bdc2e}", 100) },
            { "PropGroup_General", ("{cc301630-b192-4c22-b372-9f4c6d338e07}", 100) },
            { "PropGroup_GPS", ("{f3713ada-90e3-4e11-aae5-fdc17685b9be}", 100) },
            { "PropGroup_Image", ("{e3690a87-0fa8-4a2a-9a9f-fce8827055ac}", 100) },
            { "PropGroup_Media", ("{61872cf7-6b5e-4b4b-ac2d-59da84459248}", 100) },
            { "PropGroup_MediaAdvanced", ("{8859a284-de7e-4642-99ba-d431d044b1ec}", 100) },
            { "PropGroup_Message", ("{7fd7259d-16b4-4135-9f97-7c96ecd2fa9e}", 100) },
            { "PropGroup_Music", ("{68dd6094-7216-40f1-a029-43fe7127043f}", 100) },
            { "PropGroup_Origin", ("{2598d2fb-5569-4367-95df-5cd3a177e1a5}", 100) },
            { "PropGroup_PhotoAdvanced", ("{0cb2bf5a-9ee7-4a86-8222-f01e07fdadaf}", 100) },
            { "PropGroup_RecordedTV", ("{e7b33238-6584-4170-a5c0-ac25efd9da56}", 100) },
            { "PropGroup_Video", ("{bebe0920-7671-4c54-a3eb-49fddfc191ee}", 100) },
            { "InfoTipText", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 17) },
            { "PropList_ConflictPrompt", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 11) },
            { "PropList_ContentViewModeForBrowse", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 13) },
            { "PropList_ContentViewModeForSearch", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 14) },
            { "PropList_ExtendedTileInfo", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 9) },
            { "PropList_FileOperationPrompt", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 10) },
            { "PropList_FullDetails", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 2) },
            { "PropList_InfoTip", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 4) },
            { "PropList_NonPersonal", ("{49d1091f-082e-493f-b23f-d2308aa9668c}", 100) },
            { "PropList_PreviewDetails", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 8) },
            { "PropList_PreviewTitle", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 6) },
            { "PropList_QuickTip", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 5) },
            { "PropList_TileInfo", ("{c9944a21-a406-48fe-8225-aec7e24c211b}", 3) },
            { "PropList_XPDetailsPanel", ("{f2275480-f782-4291-bd94-f13693513aec}", 0) },
            { "RecordedTV_ChannelNumber", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 7) },
            { "RecordedTV_Credits", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 4) },
            { "RecordedTV_DateContentExpires", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 15) },
            { "RecordedTV_EpisodeName", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 2) },
            { "RecordedTV_IsATSCContent", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 16) },
            { "RecordedTV_IsClosedCaptioningAvailable", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 12) },
            { "RecordedTV_IsDTVContent", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 17) },
            { "RecordedTV_IsHDContent", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 18) },
            { "RecordedTV_IsRepeatBroadcast", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 13) },
            { "RecordedTV_IsSAP", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 14) },
            { "RecordedTV_NetworkAffiliation", ("{2c53c813-fb63-4e22-a1ab-0b331ca1e273}", 100) },
            { "RecordedTV_OriginalBroadcastDate", ("{4684fe97-8765-4842-9c13-f006447b178c}", 100) },
            { "RecordedTV_ProgramDescription", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 3) },
            { "RecordedTV_RecordingTime", ("{a5477f61-7a82-4eca-9dde-98b69b2479b3}", 100) },
            { "RecordedTV_StationCallSign", ("{6d748de2-8d38-4cc3-ac60-f009b057c557}", 5) },
            { "RecordedTV_StationName", ("{1b5439e7-eba1-4af8-bdd7-7af1d4549493}", 100) },
            { "LocationEmptyString", ("{62d2d9ab-8b64-498d-b865-402d4796f865}", 3) },
            { "Search_AutoSummary", ("{560c36c0-503a-11cf-baa1-00004c752a9a}", 2) },
            { "Search_ContainerHash", ("{bceee283-35df-4d53-826a-f36a3eefc6be}", 100) },
            { "Search_Contents", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 19) },
            { "Search_EntryID", ("{49691c90-7e17-101a-a91c-08002b2ecda9}", 5) },
            { "Search_ExtendedProperties", ("{7b03b546-fa4f-4a52-a2fe-03d5311e5865}", 100) },
            { "Search_GatherTime", ("{0b63e350-9ccc-11d0-bcdb-00805fccce04}", 8) },
            { "Search_HitCount", ("{49691c90-7e17-101a-a91c-08002b2ecda9}", 4) },
            { "Search_IsClosedDirectory", ("{0b63e343-9ccc-11d0-bcdb-00805fccce04}", 23) },
            { "Search_IsFullyContained", ("{0b63e343-9ccc-11d0-bcdb-00805fccce04}", 24) },
            { "Search_LastIndexedTotalTime", ("{0b63e350-9ccc-11d0-bcdb-00805fccce04}", 11) },
            { "Search_MatchKind", ("{49691c90-7e17-101a-a91c-08002b2ecda9}", 29) },
            { "Search_MatchTags", ("{49691c90-7e17-101a-a91c-08002b2ecda9}", 30) },
            { "Search_OcrContent", ("{b725f130-47ef-101a-a5f1-02608c9eebac}", 28) },
            { "Search_QueryFocusedSummary", ("{560c36c0-503a-11cf-baa1-00004c752a9a}", 3) },
            { "Search_QueryFocusedSummaryWithFallback", ("{560c36c0-503a-11cf-baa1-00004c752a9a}", 4) },
            { "Search_QueryPropertyHits", ("{49691c90-7e17-101a-a91c-08002b2ecda9}", 21) },
            { "Search_Rank", ("{49691c90-7e17-101a-a91c-08002b2ecda9}", 3) },
            { "Search_Store", ("{a06992b3-8caf-4ed7-a547-b259e32ac9fc}", 100) },
            { "Search_UrlToIndex", ("{0b63e343-9ccc-11d0-bcdb-00805fccce04}", 2) },
            { "Search_UrlToIndexWithModificationTime", ("{0b63e343-9ccc-11d0-bcdb-00805fccce04}", 12) },
            { "Supplemental_Album", ("{0c73b141-39d6-4653-a683-cab291eaf95b}", 6) },
            { "Supplemental_AlbumID", ("{0c73b141-39d6-4653-a683-cab291eaf95b}", 2) },
            { "Supplemental_Location", ("{0c73b141-39d6-4653-a683-cab291eaf95b}", 5) },
            { "Supplemental_Person", ("{0c73b141-39d6-4653-a683-cab291eaf95b}", 7) },
            { "Supplemental_ResourceId", ("{0c73b141-39d6-4653-a683-cab291eaf95b}", 3) },
            { "Supplemental_Tag", ("{0c73b141-39d6-4653-a683-cab291eaf95b}", 4) },
            { "ActivityInfo", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 17) },
            { "DescriptionID", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 2) },
            { "Home_Grouping", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 2) },
            { "Home_IsPinned", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 4) },
            { "Home_ItemFolderPathDisplay", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 6) },
            { "Home_RecommendationActivityDate", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 22) },
            { "Home_RecommendationProviderSource", ("{5ca9b1cb-c69f-404b-abc6-fd336793a6a7}", 22) },
            { "Home_RecommendationReasonIcon", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 21) },
            { "Home_Recommended", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 20) },
            { "InternalName", ("{0cef7d53-fa64-11d1-a203-0000f81fedee}", 5) },
            { "LibraryLocationsCount", ("{908696c7-8f87-44f2-80ed-a8c1c6894575}", 2) },
            { "Link_TargetSFGAOFlagsStrings", ("{d6942081-d53b-443d-ad47-5e059d9cd27a}", 3) },
            { "Link_TargetUrl", ("{5cbf2787-48cf-4208-b90e-ee5e5d420294}", 2) },
            { "NamespaceCLSID", ("{28636aa6-953d-11d2-b5d6-00c04fd918d0}", 6) },
            { "Shell_CopilotKeyProviderFastPathMessage", ("{38652bca-4329-4e74-86f9-39cf29345eea}", 2) },
            { "Shell_SFGAOFlagsStrings", ("{d6942081-d53b-443d-ad47-5e059d9cd27a}", 2) },
            { "StatusBarSelectedItemCount", ("{26dc287c-6e3d-4bd3-b2b0-6a26ba2e346d}", 3) },
            { "StatusBarViewItemCount", ("{26dc287c-6e3d-4bd3-b2b0-6a26ba2e346d}", 2) },
            { "StorageProviderState", ("{e77e90df-6271-4f5b-834f-2dd1f245dda4}", 3) },
            { "StorageProviderTransferProgress", ("{e77e90df-6271-4f5b-834f-2dd1f245dda4}", 4) },
            { "WebAccountID", ("{30c8eef4-a832-41e2-ab32-e3c3ca28fd29}", 7) },
            { "AppUserModel_ExcludeFromShowInNewInstall", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 8) },
            { "AppUserModel_ID", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 5) },
            { "AppUserModel_IsDestListSeparator", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 6) },
            { "AppUserModel_IsDualMode", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 11) },
            { "AppUserModel_PackageFamilyName", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 17) },
            { "AppUserModel_PreventPinning", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 9) },
            { "AppUserModel_RelaunchCommand", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 2) },
            { "AppUserModel_RelaunchDisplayNameResource", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 4) },
            { "AppUserModel_RelaunchIconResource", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 3) },
            { "AppUserModel_SettingsCommand", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 38) },
            { "AppUserModel_StartPinOption", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 12) },
            { "AppUserModel_ToastActivatorCLSID", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 26) },
            { "AppUserModel_UninstallCommand", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 37) },
            { "AppUserModel_VisualElementsManifestHintPath", ("{9f4c2855-9f79-4b39-a8d0-e1d42de1d5f3}", 31) },
            { "EdgeGesture_DisableTouchWhenFullscreen", ("{32ce38b2-2c9a-41b1-9bc5-b3784394aa44}", 2) },
            { "Software_DateLastUsed", ("{841e4f90-ff59-4d16-8947-e81bbffab36d}", 16) },
            { "Software_ProductName", ("{0cef7d53-fa64-11d1-a203-0000f81fedee}", 7) },
            { "Sync_Comments", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 13) },
            { "Sync_ConflictDescription", ("{ce50c159-2fb8-41fd-be68-d3e042e274bc}", 4) },
            { "Sync_ConflictFirstLocation", ("{ce50c159-2fb8-41fd-be68-d3e042e274bc}", 6) },
            { "Sync_ConflictSecondLocation", ("{ce50c159-2fb8-41fd-be68-d3e042e274bc}", 7) },
            { "Sync_HandlerCollectionID", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 2) },
            { "Sync_HandlerID", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 3) },
            { "Sync_HandlerName", ("{ce50c159-2fb8-41fd-be68-d3e042e274bc}", 2) },
            { "Sync_HandlerType", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 8) },
            { "Sync_HandlerTypeLabel", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 9) },
            { "Sync_ItemID", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 6) },
            { "Sync_ItemName", ("{ce50c159-2fb8-41fd-be68-d3e042e274bc}", 3) },
            { "Sync_ProgressPercentage", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 23) },
            { "Sync_State", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 24) },
            { "Sync_Status", ("{7bd5533e-af15-44db-b8c8-bd6624e1d032}", 10) },
            { "Task_BillingInformation", ("{d37d52c6-261c-4303-82b3-08b926ac6f12}", 100) },
            { "Task_CompletionStatus", ("{084d8a0a-e6d5-40de-bf1f-c8820e7c877c}", 100) },
            { "Task_Owner", ("{08c7cc5f-60f2-4494-ad75-55e3e0b5add0}", 100) },
            { "Video_Compression", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 10) },
            { "Video_Director", ("{64440492-4c8b-11d1-8b70-080036b11a03}", 20) },
            { "Video_EncodingBitrate", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 8) },
            { "Video_FourCC", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 44) },
            { "Video_FrameHeight", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 4) },
            { "Video_FrameRate", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 6) },
            { "Video_FrameWidth", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 3) },
            { "Video_HorizontalAspectRatio", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 42) },
            { "Video_IsSpherical", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 100) },
            { "Video_IsStereo", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 98) },
            { "Video_Orientation", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 99) },
            { "Video_SampleSize", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 9) },
            { "Video_StreamName", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 2) },
            { "Video_StreamNumber", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 11) },
            { "Video_TotalBitrate", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 43) },
            { "Video_TranscodedForSync", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 46) },
            { "Video_VerticalAspectRatio", ("{64440491-4c8b-11d1-8b70-080036b11a03}", 45) },
            { "Volume_FileSystem", ("{9b174b35-40ff-11d2-a27e-00c04fc30871}", 4) },
            { "Volume_IsMappedDrive", ("{149c0b69-2c2d-48fc-808f-d318d78c4636}", 2) },
            { "Volume_IsRoot", ("{9b174b35-40ff-11d2-a27e-00c04fc30871}", 10) },
        };

        private static readonly Dictionary<long, string> VariantType = new()
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
        /// https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.package.variant.varianttype?view=visualstudiosdk-2019
        /// https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum
        {
            {0x0000,"VT_EMPTY"},
            {0x0001,"VT_NULL"},
            {0x0002,"VT_I2"},
            {0x0003,"VT_I4"},
            {0x0004,"VT_R4"},
            {0x0005,"VT_R8"},
            {0x0006,"VT_CY"},
            {0x0007,"VT_DATE"},
            {0x0008,"VT_BSTR"},
            {0x0009,"VT_DISPATCH"},
            {0x000A,"VT_ERROR"},
            {0x000B,"VT_BOOL"},
            {0x000C,"VT_VARIANT"},
            {0x000D,"VT_UNKNOWN"},
            {0x000E,"VT_DECIMAL"},
            {0x0010,"VT_I1"},
            {0x0011,"VT_UI1"},
            {0x0012,"VT_UI2"},
            {0x0013,"VT_UI4"},
            {0x0014,"VT_I8"},
            {0x0015,"VT_UI8"},
            {0x0016,"VT_INT"},
            {0x0017,"VT_UINT"},
            {0x0018,"VT_VOID"},
            {0x0019,"VT_HRESULT"},
            {0x001A,"VT_PTR"},
            {0x001B,"VT_SAFEARRAY"},
            {0x001C,"VT_CARRAY"},
            {0x001D,"VT_USERDEFINED"},
            {0x001E,"VT_LPSTR"},
            {0x001F,"VT_LPWSTR"},
            {0x0024,"VT_RECORD" },
            {0x0025,"VT_INT_PTR" },
            {0x0026,"VT_UINT_PTR" },
            {0x0040,"VT_FILETIME"},
            {0x0041,"VT_BLOB"},
            {0x0042,"VT_STREAM"},
            {0x0043,"VT_STORAGE"},
            {0x0044,"VT_STREAMED_OBJECT"},
            {0x0045,"VT_STORED_OBJECT"},
            {0x0046,"VT_BLOB_OBJECT"},
            {0x0047,"VT_CF"},
            {0x0048,"VT_CLSID"},
            {0x0049,"VT_VERSIONED_STREAM"},
            {0x0FFF,"VT_BSTR_BLOB"},
            {0x1000,"VT_VECTOR"},
            {0x1002,"VT_VECTOR_I2"},
            {0x1003,"VT_VECTOR_I4"},
            {0x1004,"VT_VECTOR_R4"},
            {0x1005,"VT_VECTOR_R8"},
            {0x1006,"VT_VECTOR_CY"},
            {0x1007,"VT_VECTOR_DATE"},
            {0x1008,"VT_VECTOR_BSTR"},
            {0x100A,"VT_VECTOR_ERROR"},
            {0x100B,"VT_VECTOR_BOOL"},
            {0x100C,"VT_VECTOR_VARIANT"},
            {0x1010,"VT_VECTOR_I1"},
            {0x1011,"VT_VECTOR_UI1"},
            {0x1012,"VT_VECTOR_UI2"},
            {0x1013,"VT_VECTOR_UI4"},
            {0x1014,"VT_VECTOR_I8"},
            {0x1015,"VT_VECTOR_UI8"},
            {0x101E,"VT_VECTOR_LPSTR"},
            {0x101F,"VT_VECTOR_LPWSTR"},
            {0x1040,"VT_VECTOR_FILETIME"},
            {0x1047,"VT_VECTOR_CF"},
            {0x1048,"VT_VECTOR_CLSID"},
            {0x2000,"VT_ARRAY"},
            {0x2002,"VT_ARRAY_I2"},
            {0x2003,"VT_ARRAY_I4"},
            {0x2004,"VT_ARRAY_R4"},
            {0x2005,"VT_ARRAY_R8"},
            {0x2006,"VT_ARRAY_CY"},
            {0x2007,"VT_ARRAY_DATE"},
            {0x2008,"VT_ARRAY_BSTR"},
            {0x200A,"VT_ARRAY_ERROR"},
            {0x200B,"VT_ARRAY_BOOL"},
            {0x200C,"VT_ARRAY_VARIANT"},
            {0x200E,"VT_ARRAY_DECIMAL"},
            {0x2010,"VT_ARRAY_I1"},
            {0x2011,"VT_ARRAY_UI1"},
            {0x2012,"VT_ARRAY_UI2"},
            {0x2013,"VT_ARRAY_UI4"},
            {0x2016,"VT_ARRAY_INT"},
            {0x2017,"VT_ARRAY_UINT"},
            {0x4000,"VT_BYREF"},
            {0x8000,"VT_RESERVED"},
            {0xFFFF,"VT_ILLEGAL"}
        };

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            this.Title = $"{displayName} v{appVersion}";
            elapsedTimer!.Tick += new EventHandler(ElapsedTime!);
            elapsedTimer.Interval = TimeSpan.FromSeconds(1);
            CommandBindings.Add(new CommandBinding(KeyboardShortcuts.AboutClick, (sender, e) => { AboutClick(sender, e); }, (sender, e) => { e.CanExecute = true; }));
            InputBindings.Add(new KeyBinding(KeyboardShortcuts.AboutClick, new KeyGesture(Key.A, ModifierKeys.Control | ModifierKeys.Shift)));
            CommandBindings.Add(new CommandBinding(KeyboardShortcuts.DBPicker, (sender, e) => { DBPicker(sender, e); }, (sender, e) => { e.CanExecute = true; }));
            InputBindings.Add(new KeyBinding(KeyboardShortcuts.DBPicker, new KeyGesture(Key.D, ModifierKeys.Control)));
            CommandBindings.Add(new CommandBinding(KeyboardShortcuts.OutputPicker, (sender, e) => { OutputPicker(sender, e); }, (sender, e) => { e.CanExecute = true; }));
            InputBindings.Add(new KeyBinding(KeyboardShortcuts.OutputPicker, new KeyGesture(Key.O, ModifierKeys.Control)));
        }

        public static class KeyboardShortcuts
        {
            static KeyboardShortcuts()
            {
                AboutClick = new RoutedCommand("AboutClick", typeof(MainWindow));
                DBPicker = new RoutedCommand("DBPicker", typeof(MainWindow));
                OutputPicker = new RoutedCommand("OutputPicker", typeof(MainWindow));
            }
            public static RoutedCommand AboutClick { get; private set; }
            public static RoutedCommand DBPicker { get; private set; }
            public static RoutedCommand OutputPicker { get; private set; }

        }

        private void ElapsedTime(object source, EventArgs e)
        {
            TimeSpan timeSpan = stopWatch!.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            TimerLabel.Content = elapsedTime;
        }

        private void OutputPicker(object sender, RoutedEventArgs e)
        {
            string initialDirectory;
            if (dbFile != "")
            {
                initialDirectory = dbFile;
            }
            else
            {
                initialDirectory = Environment.SpecialFolder.Desktop.ToString();
            }
            string selectedPath = "";
            FolderBrowserDialog folderDlg = new()
            {
                Description = "Select the output directory for your processing results",
                ShowNewFolderButton = true,
                UseDescriptionForTitle = true,
                RootFolder = Environment.SpecialFolder.Desktop,
                InitialDirectory = initialDirectory
            };

            DialogResult result = folderDlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                selectedPath = folderDlg.SelectedPath;
            }
            if (selectedPath != "")
            {
                OutputPath.Text = selectedPath;
                OutputPath.ToolTip = selectedPath;
            }
        }

        private void DBPicker(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openFile = new()
                {
                    Title = "Select a Windows.db or Windows.edb file",
                    Filter = "Windows Search Index Files|*.db;*.edb;"
                };
                if (openFile.ShowDialog() == true)
                {
                    dbFile = openFile.FileName;
                    DBPath.Text = dbFile;
                    DBPath.ToolTip = dbFile;
                    OutputPath.Text = Path.GetDirectoryName(dbFile);
                    OutputPath.ToolTip = OutputPath.Text;
                }
            }
            catch (Exception ex)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Unable to load Database:\n\n{ex.Message}", "Database Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
                GoButton.IsEnabled = true;
            }
        }

        private void SelectDeselectAll(object sender, RoutedEventArgs e)
        {
            bool isSelectAll = SelectDeselectButton.Content.ToString() == "Select All";

            foreach (object child in LogicalTreeHelper.GetChildren(MainGrid))
            {
                if (child is System.Windows.Controls.CheckBox cb)
                {
                    cb.IsChecked = isSelectAll;
                }
            }

            SelectDeselectButton.Content = isSelectAll ? "Deselect All" : "Select All";
        }

        static List<string> GetMatchingFlags(long value, Dictionary<long, string> dict)
        {
            var matches = new List<string>();

            foreach (var kvp in dict)
            {
                if ((value & kvp.Key) == kvp.Key)
                {
                    matches.Add(kvp.Value);
                }
            }

            return matches;
        }

        private void ResetUI()
        {
            stopWatch?.Stop();
            elapsedTimer?.Stop();
            GoButton.IsEnabled = true;
            ProgressBarControl.IsIndeterminate = false;
            ProgressBarControl.Value = 100;
            ProgressBarControl.Visibility = Visibility.Collapsed;
            StatusBox.Visibility = Visibility.Collapsed;
            StatusBox.Text = "";
            StatusBox.Foreground = System.Windows.Media.Brushes.Black;
            TimerLabel.Visibility = Visibility.Collapsed;
            URLCheck.IsEnabled = true;
            URLCheck.IsChecked = false;
            GPSCheck.IsEnabled = true;
            GPSCheck.IsChecked = false;
            SummaryCheck.IsEnabled = true;
            SummaryCheck.IsChecked = false;
            CompInfoCheck.IsEnabled = true;
            CompInfoCheck.IsChecked = false;
            ActivityCheck.IsEnabled = true;
            ActivityCheck.IsChecked = false;
            TimelineCheck.IsEnabled = true;
            TimelineCheck.IsChecked = false;
            DBPath.IsEnabled = true;
            DBPathPicker.IsEnabled = true;
            OutputPath.IsEnabled = true;
            OutputPathPicker.IsEnabled = true;
            SelectDeselectButton.Content = "Select All";
            SelectDeselectButton.IsEnabled = true;
            //IndexProperties.Clear();
            IndexResults.Clear();
            GatherResults.Clear();
            URLResults.Clear();
            ActivityResults.Clear();
            SummaryResults.Clear();
            GPSResults.Clear();
            CompInfoResults.Clear();
            Timeline.Clear();
            properties.Clear();
            EseResults.Clear();
            propertyStore.Clear();
            propertyMetadata.Clear();
            tables.Clear();
        }

        private void SetupUI()
        {
            stopWatch?.Reset();
            stopWatch?.Start();
            elapsedTimer?.Start();
            GoButton.IsEnabled = false;
            ProgressBarControl.Visibility = Visibility.Visible;
            StatusBox.Visibility = Visibility.Visible;
            TimerLabel.Visibility = Visibility.Visible;
            ProgressBarControl.IsIndeterminate = true;
            URLCheck.IsEnabled = false;
            GPSCheck.IsEnabled = false;
            SummaryCheck.IsEnabled = false;
            CompInfoCheck.IsEnabled = false;
            ActivityCheck.IsEnabled = false;
            TimelineCheck.IsEnabled = false;
            SelectDeselectButton.IsEnabled = false;
            DBPath.IsEnabled = false;
            DBPathPicker.IsEnabled = false;
            OutputPath.IsEnabled = false;
            OutputPathPicker.IsEnabled = false;
            rows.Clear();
            paths.Clear();
            //IndexProperties.Clear();
            IndexResults.Clear();
            GatherResults.Clear();
            URLResults.Clear();
            ActivityResults.Clear();
            SummaryResults.Clear();
            GPSResults.Clear();
            CompInfoResults.Clear();
            Timeline.Clear();
            properties.Clear();
            EseResults.Clear();
            propertyStore.Clear();
            propertyMetadata.Clear();
            tables.Clear();
        }

        private async void GoButtonClick(object sender, RoutedEventArgs e)
        {
            DateTime nowDt = DateTime.Now.ToUniversalTime();
            string now = nowDt.ToString("yyyyMMdd-HHmmss");
            try
            {
                if (DBPath.Text == "")
                {
                    App.Current.MainWindow.Activate();
                    System.Windows.MessageBox.Show("No database was selected.\n\nPlease select a valid database path and try again.", "Database Path is required.", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                if (OutputPath.Text == "")
                {
                    App.Current.MainWindow.Activate();
                    System.Windows.MessageBox.Show("No output path was selected.\n\nPlease select a valid output path and try again.", "Output Path is required.", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                SetupUI();
                await Task.Yield();
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
                StatusBox.Text = "Extracting data ...";
                await Task.Run(() => ExtractData());
                if (!dbFile.Contains("gather") && dbType == "index")
                {
                    GatherAvailable = File.Exists(dbFile.Replace("Windows.db", "Windows-gather.db"));
                    if (GatherAvailable)
                    {
                        MessageBoxResult result = System.Windows.MessageBox.Show("Windexter has found a Windows-gather.db in the same directory as the selected database.\n\nDo you want to process this as well?", "Windows-gather.db exists! Process?", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                        if (result == MessageBoxResult.Yes)
                        {
                            dbFile = dbFile.Replace("Windows.db", "Windows-gather.db");
                            await Task.Run(() => ExtractData());
                            dbType = "index";
                        }
                        else if (result == MessageBoxResult.Cancel)
                        {
                            ResetUI();
                            return;
                        }
                        else if (result == MessageBoxResult.No)
                        {
                            GatherAvailable = false;
                        }
                    }
                    string[] currentDirFiles = Directory.GetFiles(Path.GetDirectoryName(dbFile)!, "PropMap.db", SearchOption.AllDirectories);
                    var propPath = "";
                    if (currentDirFiles.Length > 0)
                    {
                        propPath = currentDirFiles[0];
                        PropMapAvailable = true;
                    }
                    if (PropMapAvailable)
                    {
                        MessageBoxResult result = System.Windows.MessageBox.Show("Windexter has found a PropMap.db in a sub-directory of the selected folder.\n\nDo you want to process this as well?", "PropMap.db exists! Process?", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                        if (result == MessageBoxResult.Yes)
                        {
                            dbFile = propPath;
                            await Task.Run(() => ExtractData());
                            dbType = "index";
                        }
                        else if (result == MessageBoxResult.Cancel)
                        {
                            ResetUI();
                            return;
                        }
                        else if (result == MessageBoxResult.No)
                        {
                            PropMapAvailable = false;
                        }
                    }
                }
                outputFile = Path.Combine(OutputPath.Text, $"WINDOWS-SEARCH-{dbType.ToUpper()}-{now}.xlsx");
                StatusBox.Text = "Parsing data ...";
                await ParseDataAsync();
                StatusBox.Text = "Exporting to XLSX...";
                bool cbsChecked = AnyCheckBoxChecked(MainGrid);
                if (cbsChecked)
                {
                    await ParseSubData();
                }
                var AllResults = new Dictionary<string, List<List<object>>>
                {
                    ["Indexed Results"] = IndexResults,
                    //["Index Properties"] = IndexProperties,
                    ["Gather Data"] = GatherResults,
                    ["URL Data"] = URLResults,
                    ["GPS Data"] = GPSResults,
                    ["Computer Info"] = CompInfoResults,
                    ["Activity"] = ActivityResults,
                    ["Search Summary"] = SummaryResults,
                    ["Timeline"] = Timeline,
                    ["Property Map"] = PropertyMapResults,
                };
                if (!await ExportToExcel(outputFile, AllResults))
                {
                    return;
                }
                StatusBox.Text = "Extraction and parsing completed!";
                ProgressBarControl.Value = 100;
                ProgressBarControl.IsIndeterminate = false;
                StatusBox.Foreground = System.Windows.Media.Brushes.White;
                stopWatch?.Stop();
                elapsedTimer?.Stop();
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Data successfully exported to:\n\n{outputFile}", "Export Successful", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                stopWatch?.Stop();
                elapsedTimer?.Stop();
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Error with data extraction:\n\n{ex.Message}\n{ex}", "Data Extraction Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ResetUI();
            }
        }

        private static bool AnyCheckBoxChecked(DependencyObject parent)
        {
            foreach (object child in LogicalTreeHelper.GetChildren(parent))
            {
                if (child is System.Windows.Controls.CheckBox cb && cb.IsChecked == true)
                {
                    return true;
                }

                if (child is DependencyObject d)
                {
                    if (AnyCheckBoxChecked(d)) return true;
                }
            }
            return false;
        }

        private async Task ParseDataAsync()
        {
            DateTime nowDt = DateTime.Now.ToUniversalTime();
            string now = nowDt.ToString("yyyyMMdd-HHmmss");
            if (dbType == "index")
            {
                if (propertyMetadata.Count > 0 && propertyStore.Count > 0)
                {
                    try
                    {
                        await Task.Run(() => GetIndexPropertyStore(propertyStore, propertyMetadata));
                        if (GatherAvailable)
                        {
                            ParseGatherData();
                            GatherAvailable = false;
                        }
                        if (PropMapAvailable)
                        {
                            ParsePropertyMap();
                            PropMapAvailable = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        App.Current.MainWindow.Activate();
                        System.Windows.MessageBox.Show($"Unable to parse Index database:\n\n{ex.Message}", "Index Database Parsing Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        GoButton.IsEnabled = true;
                    }
                }
            }
            else if (dbType == "esedb" || dbType == "gather")
            {
                try
                {
                    ParseGatherData();
                    if (dbType == "esedb")
                    {
                        GetEsePropertyStore();
                        //GetEseProperties();
                    }
                }
                catch (Exception ex)
                {
                    App.Current.MainWindow.Activate();
                    System.Windows.MessageBox.Show($"Unable to parse Gather Data:\n\n{ex.Message}", "Gather Data Parsing Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    GoButton.IsEnabled = true;
                }
            }
        }

        private void ParseGatherData()
        {
            DateTime nowDt = DateTime.Now.ToUniversalTime();
            string now = nowDt.ToString("yyyyMMdd-HHmmss");
            try
            {
                var resolvedPathsDict = resolvedPaths.ToDictionary(kvp => (long)kvp.Key, kvp => kvp.Value);
                List<object> header = rows[0];
                int fn = header.IndexOf("FileName");
                List<object> newHeader = ["FullPath"];
                newHeader.AddRange(header);
                GatherResults.Add(newHeader);
                for (int i = 1; i < rows.Count; i++)
                {
                    var row = rows[i];
                    if (row.Count > 0)
                    {
                        var newRow = new List<object>();
                        string fullPath = "";
                        for (int j = 0; j < row.Count; j++)
                        {
                            object item = row[j];
                            object value = item;
                            string colName = header[j].ToString()!;
                            if (colName == "ScopeID")
                            {
                                object key = dbType == "esedb" ? (uint)item : (long)item;
                                if (row[fn] is string fileName)
                                {
                                    string path = resolvedPathsDict.GetValueOrDefault((long)key, "::Unknown::/");
                                    fullPath = string.Concat(path, fileName);
                                    newRow.Insert(0, fullPath);
                                    value = key;
                                }
                            }
                            else if (colName == "UserData" && item is byte[] userDataBytes)
                            {
                                value = SPSParser.ParseToJson(userDataBytes);
                            }
                            else if (colName == "LastModified" && item is byte[] dtcellBytes)
                            {
                                if (dtcellBytes.Length == 8 && !dtcellBytes.All(b => b == 0x2A))
                                {
                                    if (dbType == "esedb")
                                    {
                                        Array.Reverse(dtcellBytes);
                                    }
                                    long filetime = BitConverter.ToInt64(dtcellBytes, 0);
                                    if (filetime != 1)
                                    {
                                        value = DateTime.FromFileTimeUtc(filetime).ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    else
                                    {
                                        value = null!;
                                    }
                                }
                                else
                                {
                                    value = null!;
                                }
                            }
                            else if (colName == "StartAddressIdentifier" && item is byte[] startAddId)
                            {
                                value = BitConverter.ToInt16(startAddId, 0);
                            }
                            else if ((colName == "Priority" || colName == "FailureUpdateAttempts") && item is byte[] singleByteArr)
                            {
                                value = singleByteArr[0];
                            }
                            newRow.Add(value);
                        }
                        GatherResults.Add(newRow);
                    }
                }
            }
            catch (Exception ex)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Unable to parse Gather Data:\n\n{ex.Message}", "Gather Data Parsing Error", MessageBoxButton.OK, MessageBoxImage.Error);
                GoButton.IsEnabled = true;
            }
        }

        private static void ParsePropertyMap()
        {
            List<object> header = propertyMap[0];
            PropertyMapResults.Add(["Guid", "FormatId", "PropertyId", "StandardId", "MaxSize"]);
            for (int i = 1; i < propertyMap.Count; i++)
            {
                var row = propertyMap[i];
                if (row.Count > 0)
                {
                    var newRow = new List<object>();
                    for (int j = 0; j < row.Count; j++)
                    {
                        object item = row[j];
                        object value = item;
                        string colName = header[j].ToString()!;
                        if (colName == "FormatId")
                        {
                            value = BitConverter.ToString((byte[])value).Replace("-","");
                            Guid guid = new((string)value);
                            value = guid.ToString("B").ToUpper();
                        }
                        newRow.Add(value);
                    }
                    PropertyMapResults.Add(newRow);
                }
            }
        }

        private static async Task GenerateTimelineAsync()
        {
            var result = await Task.Run(() =>
            {
                var header = IndexResults[0].Cast<string>().ToList();

                var dateTimeIndexes = header
                    .Select((name, index) => (name, index))
                    .Where(x => dateTimes.Contains(x.name))
                    .ToList();

                var staticIndexes = header
                    .Select((name, index) => (name, index))
                    .Where(x => !dateTimes.Contains(x.name))
                    .ToList();
                var tempTimeline = new List<List<object>>();
                var timelineHeader = new List<object>
            {
                "Timestamp",
                "Source"
            };
                timelineHeader.AddRange(staticIndexes.Select(x => x.name));
                tempTimeline.Add(timelineHeader);
                for (int r = 1; r < IndexResults.Count; r++)
                {
                    var row = IndexResults[r];

                    foreach (var (sourceName, dtIndex) in dateTimeIndexes)
                    {
                        var value = row[dtIndex];

                        if (value is not string timestamp)
                            continue;

                        var timelineRow = new List<object>
                    {
                        timestamp,
                        sourceName
                    };
                        foreach (var (_, staticIndex) in staticIndexes)
                        {
                            timelineRow.Add(row[staticIndex]);
                        }
                        tempTimeline.Add(timelineRow);
                    }
                }
                return tempTimeline;
            });
            Timeline = result;
        }

        public async Task ParseSubData()
        {
            if (dbType == "index" || dbType == "esedb")
            {
                if ((bool)URLCheck.IsChecked!)
                {
                    var selectedColumns = new List<string>
                    { 
                        "WorkId", 
                        "System.Search.GatherTime", 
                        "System.DateCreated", 
                        "System.DateModified", 
                        "System.ItemType", 
                        "System.Link.TargetUrl", 
                        "System.History.VisitCount"
                    };
                    List<List<object>> temp = FilterColumns(IndexResults, selectedColumns);
                    int urlIdx = temp[0].IndexOf("System.Link.TargetUrl");
                    if (urlIdx > 0)
                    {
                        URLResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if (row[urlIdx] is string s && !string.IsNullOrEmpty(s))
                            {
                                URLResults.Add(row);
                            }
                        }
                    }
                }
                if ((bool)GPSCheck.IsChecked!)
                {
                    var selectedColumns = new List<string>
                    { 
                        "WorkId", 
                        "System.Search.GatherTime", 
                        "System.ComputerName", 
                        "System.VolumeId", 
                        "System.FileName", 
                        "System.ItemPathDisplay", 
                        "System.DateCreated", 
                        "System.DateModified", 
                        "System.GPS.LatitudeRef", 
                        "System.GPS.LatitudeDecimal", 
                        "System.GPS.LongitudeRef", 
                        "System.GPS.LongitudeDecimal", 
                        "System.Photo.DateTaken", 
                        "System.Photo.CameraManufacturer", 
                        "System.Photo.CameraModel" 
                    };
                    List<List<object>> temp = FilterColumns(IndexResults, selectedColumns);
                    int latIdx = temp[0].IndexOf("System.GPS.LatitudeDecimal");
                    int longIdx = temp[0].IndexOf("System.GPS.LongitudeDecimal");
                    if (latIdx > 0 && longIdx > 0)
                    {
                        GPSResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if ((row[latIdx] is double lat && !double.IsNaN(lat)) || (row[longIdx] is double lon && !double.IsNaN(lon)))
                            {
                                GPSResults.Add(row);
                            }
                        }
                    }
                }
                if ((bool)SummaryCheck.IsChecked!)
                {
                    var selectedColumns = new List<string>
                    { 
                        "WorkId", 
                        "System.Search.GatherTime", 
                        "System.DateCreated", 
                        "System.DateModified", 
                        "System.DateAccessed", 
                        "System.ItemPathDisplay", 
                        "System.FileName", 
                        "System.Size", 
                        "System.Search.AutoSummary", 
                        "System.KindText", 
                        "System.IsDeleted", 
                        "System.ItemType" 
                    };
                    List<List<object>> temp = FilterColumns(IndexResults, selectedColumns);
                    int summaryIdx = temp[0].IndexOf("System.Search.AutoSummary");
                    if (summaryIdx > 0)
                    {
                        SummaryResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if (row[summaryIdx] is string s && !string.IsNullOrEmpty(s))
                            {
                                SummaryResults.Add(row);
                            }
                        }
                    }
                }
                if ((bool)CompInfoCheck.IsChecked!)
                {
                    var selectedColumns = new List<string>
                    { 
                        "WorkId", 
                        "System.Search.GatherTime", 
                        "System.VolumeId", 
                        "System.ComputerName", 
                        "System.DateCreated", 
                        "System.DateModified", 
                        "System.DateAccessed", 
                        "System.ItemType" 
                    };
                    List<List<object>> temp = FilterColumns(IndexResults, selectedColumns);
                    int itemTypeIdx = temp[0].IndexOf("System.ItemType");
                    if (itemTypeIdx > 0)
                    {
                        CompInfoResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        List<string> itemTypes = [".library-ms", ".searchconnector-ms", ".search-ms"];
                        foreach (List<object> row in temp)
                        {
                            if (row[itemTypeIdx] is string s && itemTypes.Contains(s))
                            {
                                CompInfoResults.Add(row);
                            }
                        }
                    }
                }
                if ((bool)ActivityCheck.IsChecked!)
                {
                    var selectedColumns = new List<string>
                    { 
                        "WorkId", 
                        "System.Search.GatherTime", 
                        "System.ComputerName", 
                        "System.VolumeId", 
                        "System.DateModified", 
                        "System.ActivityHistory.StartTime", 
                        "System.ActivityHistory.EndTime", 
                        "System.ItemNameDisplay", 
                        "System.ItemUrl", 
                        "System.Activity.AppDisplayName", 
                        "System.Activity.ContentUri", 
                        "System.Activity.DisplayText", 
                        "System.Activity.Description", 
                        "System.ActivityHistory.AppId", 
                        "System.ActivityHistory.AppIdList", 
                        "System.ItemType" 
                    };
                    List<List<object>> temp = FilterColumns(IndexResults, selectedColumns);
                    int itemTypeIdx = temp[0].IndexOf("System.ItemType");
                    if (itemTypeIdx > 0)
                    {
                        ActivityResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if (row[itemTypeIdx] is string s && s == "ActivityHistoryItem")
                            {
                                ActivityResults.Add(row);
                            }
                        }
                    }
                }
                if ((bool)TimelineCheck.IsChecked!)
                {
                    StatusBox.Text = "Generating timeline ...";
                    await GenerateTimelineAsync();
                }
            }
        }

        private void ExtractData()
        {
            if (dbFile.EndsWith(".db"))
            {
                Batteries_V2.Init();
                sqlite3 db = null!;
                sqlite3_stmt stm = null!;
                sqlite3_stmt stmt_path = null!;
                sqlite3_stmt cat_stmt = null!;

                try
                {
                    int rc = raw.sqlite3_open_v2(dbFile, out db, raw.SQLITE_OPEN_READONLY, null);
                    if (rc != raw.SQLITE_OK)
                    {
                        throw new Exception($"Failed to open database: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                    }

                    rc = raw.sqlite3_create_collation(db, "UNICODE_en-US_LINGUISTIC_IGNORECASE", null,
                        (userData, a, b) => string.Compare(a, b, StringComparison.OrdinalIgnoreCase));
                    if (rc != raw.SQLITE_OK)
                    {
                        throw new Exception($"Failed to register collation: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                    }

                    string sql = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';";
                    rc = raw.sqlite3_prepare_v2(db, sql, out stm);
                    if (rc != raw.SQLITE_OK)
                    {
                        throw new Exception($"Failed to prepare SQL statement: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                    }

                    while (raw.sqlite3_step(stm) == raw.SQLITE_ROW)
                    {
                        string? tableName = raw.sqlite3_column_text(stm, 0).utf8_to_string();
                        if (tableName != null)
                        {
                            tables.Add(tableName);
                        }
                    }
                    raw.sqlite3_finalize(stm);
                    stm = null!;
                    if (tables.Contains("SystemIndex_Gthr"))
                    {
                        dbType = "gather";
                        sql = "SELECT * FROM SystemIndex_Gthr";
                        rows = ReadAndStoreTable(db, sql);

                        string sql_path = "SELECT * FROM SystemIndex_GthrPth";
                        rc = raw.sqlite3_prepare_v2(db, sql_path, out stmt_path);
                        if (rc != raw.SQLITE_OK)
                        {
                            throw new Exception($"Failed to prepare SELECT * query for SystemIndex_GthrPth: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                        }
                        while ((rc = raw.sqlite3_step(stmt_path)) == raw.SQLITE_ROW)
                        {
                            int trhColCount = raw.sqlite3_column_count(stmt_path);
                            List<string> cols = new(trhColCount);
                            for (int i = 0; i < trhColCount; i++)
                            {
                                string? text = raw.sqlite3_column_text(stmt_path, i).utf8_to_string();
                                cols.Add(text ?? string.Empty);
                            }
                            paths.Add(cols);
                        }
                        if (rc != raw.SQLITE_DONE)
                        {
                            throw new Exception($"Error iterating SystemIndex_GthrPth rows: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                        }
                        raw.sqlite3_finalize(stmt_path);
                        stmt_path = null!;
                        resolvedPaths = BuildFullPaths(paths);
                    }
                    else if (tables.Contains("CatalogStorageManager"))
                    {
                        dbType = "index";
                        string propStoreTable = tables.FirstOrDefault(t => t.EndsWith("_PropertyStore")) ?? string.Empty;
                        string metaTable = tables.FirstOrDefault(t => t.EndsWith("_Metadata")) ?? string.Empty;
                        string propTable = tables.FirstOrDefault(t => t.EndsWith("_Properties")) ?? string.Empty;

                        if (string.IsNullOrEmpty(propStoreTable) || string.IsNullOrEmpty(metaTable))
                        {
                            return;
                        }

                        sql = $"SELECT * FROM {propStoreTable}";
                        propertyStore = ReadAndStoreTable(db, sql);

                        sql = $"SELECT * FROM {metaTable}";
                        propertyMetadata = ReadAndStoreTable(db, sql);

                        sql = $"SELECT * FROM {propTable}";
                        properties = ReadAndStoreTable(db, sql);
                    }
                    else if (tables.Contains("PropertyMap"))
                    {
                        sql = $"SELECT * FROM PropertyMap";
                        propertyMap = ReadAndStoreTable(db, sql);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to extract data:\n{ex}");
                }
                finally
                {
                    if (stm != null) raw.sqlite3_finalize(stm);
                    if (stmt_path != null) raw.sqlite3_finalize(stmt_path);
                    if (cat_stmt != null) raw.sqlite3_finalize(cat_stmt);
                    if (db != null) raw.sqlite3_close(db);
                    tables.Clear();
                }
            }
            else if (dbFile.EndsWith(".edb"))
            {
                dbType = "esedb";
                try
                {
                    rows.Clear();
                    using var reader = new LibEsedbReader();
                    reader.Open(dbFile);
                    var allData = reader.ReadAllData();
                    var gatherPaths = allData["SystemIndex_GthrPth"];
                    var gatherData = allData["SystemIndex_Gthr"];
                    rows.Add([.. ((IEnumerable)gatherData[0]).Cast<object>()]);
                    foreach (List<object> row in (IEnumerable)gatherData.Skip(1))
                    {
                        rows.Add(row);
                    }
                    resolvedPaths = BuildEsePaths(gatherPaths);
                    var propStore = allData["SystemIndex_PropertyStore"];
                    var propStoreHeader = (IEnumerable)propStore[0];
                    List<string> newHeader = [];
                    foreach (string title in propStoreHeader)
                    {
                        string newTitle;
                        if (title.Contains('-'))
                        {
                            string[] split = title.Split('-', 2);
                            var key = split[0];
                            string value = split[1].Replace("_",".");
                            newTitle = value;
                            eseProps[key] = value;
                        }
                        else
                        {
                            newTitle = title.Replace("_", ".");
                        }
                        newHeader.Add(newTitle);
                    }
                    propertyStore.Add([.. newHeader.Cast<object>()]);
                    foreach (List<object> row in (IEnumerable)propStore.Skip(1))
                    {
                        propertyStore.Add(row);
                    }
                    var eseProperties = allData["SystemIndex_1_Properties"];
                    properties.Add([.. ((IEnumerable)eseProperties[0]).Cast<object>()]);
                    foreach (List<object> row in (IEnumerable)eseProperties.Skip(1))
                    {
                        properties.Add(row);
                    }
                    tables.Clear();
                }
                catch (Exception ex)
                {
                    App.Current.MainWindow.Activate();
                    System.Windows.MessageBox.Show($"Unable to extract data from ESE database:\n\n{ex.Message}", "Database Parsing Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    GoButton.IsEnabled = true;
                }
            }
        }

        public static Dictionary<int, string> BuildEsePaths(List<object> gatherPaths)
        {
            List<List<string>> paths = [];
            foreach (List<string> listEntry in gatherPaths.Skip(1).Cast<List<string>>())
            {
                paths.Add(listEntry);
            }
            var builtPaths = BuildFullPaths(paths);
            return builtPaths;
        }

        private static List<List<object>> ReadAndStoreTable(sqlite3 db, string sql)
        {
            var results = new List<List<object>>();
            int rc = raw.sqlite3_prepare_v2(db, sql, out sqlite3_stmt stmt);
            if (rc != raw.SQLITE_OK)
            {
                throw new Exception($"Failed to prepare SQL statement: {raw.sqlite3_errmsg(db).utf8_to_string()}");
            }

            try
            {
                int colCount = raw.sqlite3_column_count(stmt);

                var headers = new List<object>(colCount);
                for (int i = 0; i < colCount; i++)
                {
                    string? headerName = raw.sqlite3_column_name(stmt, i).utf8_to_string();
                    headers.Add(headerName ?? string.Empty);
                }
                results.Add(headers);

                while ((rc = raw.sqlite3_step(stmt)) == raw.SQLITE_ROW)
                {
                    var colValues = new List<object>(colCount);
                    for (int i = 0; i < colCount; i++)
                    {
                        int colType = raw.sqlite3_column_type(stmt, i);
                        object value;
                        switch (colType)
                        {
                            case raw.SQLITE_INTEGER:
                                value = raw.sqlite3_column_int64(stmt, i);
                                break;
                            case raw.SQLITE_FLOAT:
                                value = raw.sqlite3_column_double(stmt, i);
                                break;
                            case raw.SQLITE_TEXT:
                                value = raw.sqlite3_column_text(stmt, i).utf8_to_string();
                                break;
                            case raw.SQLITE_NULL:
                                value = DBNull.Value;
                                break;
                            case raw.SQLITE_BLOB:
                                {
                                    ReadOnlySpan<byte> blobSpan = raw.sqlite3_column_blob(stmt, i);
                                    value = blobSpan.ToArray();
                                    break;
                                }
                            default:
                                throw new NotSupportedException($"Unsupported column type {colType} at column {i}");
                        }
                        colValues.Add(value);
                    }
                    results.Add(colValues);
                }
                if (rc != raw.SQLITE_DONE)
                {
                    throw new Exception($"Error iterating rows: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                }
            }
            catch (Exception ex)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Unable to Read and Store data:\n\n{ex.Message}", "Database Read and Store Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (stmt != null)
                {
                    raw.sqlite3_finalize(stmt);
                }
            }
            return results;
        }

        private static void GetEseProperties()
        {
            var header = properties[0];
            //IndexProperties = [header];
            foreach (var entry in properties.Skip(1))
            {
                var row = new List<object> { };
                for (int j = 0; j < entry.Count; j++)
                {
                    object item = entry[j];
                    string colName = header[j].ToString()!;
                    object value = item;
                    if (colName == "Property" && value is byte[] bytes)
                    {
                        value = BitConverter.ToString(bytes).Replace("-", "");
                    }
                    row.Add(value);
                }
                //IndexProperties.Add(row);
            }
        }

        private void GetEsePropertyStore()
        {
            try
            {
                DateTime nowDt = DateTime.Now.ToUniversalTime();
                DateTime futureDt = nowDt.AddYears(5);
                DateTime ftEpoch = new(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                long futureTicks = (futureDt - ftEpoch).Ticks;
                var header = propertyStore[0];
                IndexResults = [header];
                foreach (var entry in propertyStore.Skip(1))
                {
                    var row = new List<object> { };
                    for (int j = 0; j < entry.Count; j++)
                    {
                        object item = entry[j];
                        string colName = header[j].ToString()!;
                        object value = item;
                        if (booleanField.Contains(colName))
                        {
                            if (value is byte[] boolBytes)
                            {
                                value = BitConverter.ToBoolean(boolBytes);
                            }
                        }
                        if (value is byte[] bytes)
                        {
                            if (bytes.Length == 8 && bytes.All(b => b == 0x2A))
                            {
                                value = null!;
                            }
                            else if (bytes.Length == 8 && dateTimes.Contains(colName.ToString()!))
                            {
                                try
                                {
                                    long fileTime = BitConverter.ToInt64(bytes, 0);
                                    if (fileTime <= futureTicks && fileTime > 116444736000000000)
                                    {
                                        value = DateTime.FromFileTimeUtc(fileTime).ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    else
                                    {
                                        value = "";
                                    }
                                }
                                catch
                                {
                                    value = BitConverter.ToString(bytes).Replace("-", "");
                                }
                            }
                            else if (unicodeField.Contains(colName.ToString()!))
                            {
                                if (bytes[1] != 0)
                                {
                                    value = Decompress7Bit(bytes).Replace('\0',' ').Trim();
                                }
                                else
                                {
                                    value = Encoding.Unicode.GetString(bytes).Replace('\0', ' ').Trim();
                                }
                            }
                            else if (uint64List.Contains(colName.ToString()!))
                            {
                                value = BitConverter.ToUInt64(bytes);
                            }
                            else if (colName.ToString()!.Equals("System.Size", StringComparison.OrdinalIgnoreCase))
                            {
                                long fileSize = BitConverter.ToInt64(bytes, 0);
                                value = fileSize;

                            }
                            else if (colName.ToString()!.Equals("InvertedOnlyPids", StringComparison.OrdinalIgnoreCase))
                            {
                                List<string> pids = [];
                                for (int i = 0; i < bytes.Length; i += 2)
                                {
                                    string pid = BitConverter.ToInt16(bytes, i).ToString();
                                    pids.Add(pid);
                                    
                                }
                                value = string.Join("|", pids);
                            }
                            else if (colName.ToString()!.Equals("System.Message.ConversationIndex", StringComparison.OrdinalIgnoreCase))
                            {
                                List<string> indices = [];
                                for (int i = 0; i < bytes.Length; i += 2)
                                {
                                    string index = BitConverter.ToUInt16(bytes, i).ToString();
                                    indices.Add(index);
                                }
                                value = string.Join("|", indices);
                            }
                            else if (durations.Contains(colName.ToString()!))
                            {
                                ulong time = BitConverter.ToUInt64((byte[])value, 0);
                                TimeSpan duration = TimeSpan.FromTicks((long)time);
                                value = string.Format("{0:D2}:{1:D2}:{2:D2}:{3:D2}", (int)duration.Days, (int)duration.Hours, duration.Minutes, duration.Seconds);
                                var description = GetMeasurementName(time, MediaDurations);
                                value = $"{value} - {description}";
                            }
                            else if (floatField.Contains(colName.ToString()!))
                            {
                                value = BitConverter.ToDouble(bytes);
                            }
                            else if (colName.ToString()!.Equals("System.GPS.Longitude") | colName.ToString()!.Equals("System.GPS.Latitude"))
                            {
                                value = BitConverter.ToDouble(bytes);
                            }
                            else if (guids.Contains(colName.ToString()!))
                            {
                                value = BitConverter.ToString(bytes).Replace("-", "");
                                Guid guid = new((string)value);
                                value = guid.ToString("B").ToUpper();
                            }
                            else if (lookups.Contains(colName.ToString()!) && value is not null)
                            {
                                if (LookupValues.TryGetValue(colName.ToString()!, out var lookup))
                                {
                                    value = BitConverter.ToInt16(bytes, 0);
                                    if (lookup.TryGetValue((short)value, out string? strValue))
                                    {
                                        value = strValue;
                                    }
                                }
                            }
                            else if (bytesToString.Contains(colName.ToString()!))
                            {
                                value = BitConverter.ToString(bytes).Replace("-", "");
                            }
                            else
                            {
                                value = BitConverter.ToString(bytes).Replace("-", "");
                            }
                        }
                        if (sfgaoField.Contains(colName.ToString()!) && value is not null)
                        {
                            List<string> flags = GetMatchingFlags((uint)value, SFGAO);
                            value = string.Join("|", flags);
                        }
                        if (colName.ToString()!.Equals("System.Activity.BackgroundColor", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            uint colorVal = (uint)value;
                            Color c = Color.FromArgb((int)colorVal);
                            value = $"{value} - ARGB({c.A},{c.R},{c.G},{c.B})";
                        }
                        if (colName.ToString()!.Equals("System.FilePlaceholderStatus", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            List<string> matches = GetMatchingFlags((uint)value, FilePlaceholderStates);
                            value = string.Join("|", matches);
                        }
                        if (colName.ToString()!.Equals("System.FileAttributes", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            List<string> flags = GetMatchingFlags((uint)value, FileAttributes);
                            value = string.Join("|", flags);
                        }
                        if (colName.ToString()!.Equals("System.Capacity", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            string capacityName = GetMeasurementName((ulong)value, Capacity);
                            value = (string)value + capacityName;
                        }
                        if (colName.ToString()!.Equals("System.Message.Flags", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            List<string> matches = GetMatchingFlags((uint)value, MessageFlags);
                            value = string.Join("|", matches);
                        }
                        row.Add(value!);
                    }
                    IndexResults.Add(row);
                }

            }
            catch (Exception ex)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Unable to read and correlate data:\n\n{ex.Message}", "Data Correlation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                GoButton.IsEnabled = true;
            }
        }

        public static string Decompress7Bit(byte[] data)
        /// https://www.3gpp.org/ftp/Specs/archive/23_series/23.038 Para 6, SMS Packing
        /// https://doubleblak.com/blogPost.php?k=7bitpdu
        {
            if (data == null || data.Length <= 1)
                return string.Empty;
            StringBuilder sb = new();
            int shift = 0;
            int prevByte = 0;
            for (int i = 1; i < data.Length; i++)
            {
                byte currentByte = data[i];
                int charCode = ((currentByte << shift) & 0x7F) | (prevByte >> (8 - shift));
                sb.Append((char)charCode);

                shift++;
                if (shift == 7)
                {
                    // After every 7 bytes processed, there are 7 "leftover" bits in the current byte.
                    // This forms an 8th character.
                    int eighthChar = (currentByte >> 1) & 0x7F;
                    sb.Append((char)eighthChar);
                    shift = 0;
                    prevByte = 0;
                }
                else
                {
                    prevByte = currentByte;
                }
            }
            return sb.ToString();
        }

        private void GetIndexPropertyStore(List<List<object>> propertyStore, List<List<object>> propertyMetadata)
        {
            try
            {
                DateTime nowDt = DateTime.Now.ToUniversalTime();
                DateTime futureDt = nowDt.AddYears(5);
                DateTime ftEpoch = new(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                long futureTicks = (futureDt - ftEpoch).Ticks;
                var idToName = propertyMetadata.Skip(1).ToDictionary(row => (long)row[0], row => (string)row[2]);
                var pivotedData = propertyStore.Skip(1).GroupBy(row => (long)row[0]).ToDictionary(group => group.Key, group => group.ToDictionary(item => (long)item[1], item => item[2]));
                var workIds = pivotedData.Keys.ToList();
                var uniquePropIds = propertyStore.Skip(1).Select(row => (long)row[1]).Distinct().ToList();
                var header = new List<object> { "WorkId" };
                var seenPropNames = new HashSet<string>();
                foreach (var propId in uniquePropIds)
                {
                    if (idToName.TryGetValue(propId, out var propName) && !seenPropNames.Contains(propName))
                    {
                        header.Add(propName);
                        seenPropNames.Add(propName);
                    }
                }
                IndexResults = [header];
                foreach (var workId in workIds)
                {
                    var row = new List<object> { workId };
                    var workIdProperties = pivotedData[workId];
                    foreach (var colName in header.Skip(1))
                    {
                        var colId = idToName.FirstOrDefault(kvp => kvp.Value == colName.ToString()).Key;
                        if (colId != 0 && workIdProperties.TryGetValue(colId, out object? value))
                        {
                            if (booleanField.Contains(colName.ToString()!))
                            {
                                value = Convert.ToBoolean(value);
                            }
                            if (value is byte[] bytes)
                            {
                                if (bytes.Length == 8 && bytes.All(b => b == 0x2A))
                                {
                                    value = null!;
                                }
                                else if (bytes.Length == 8 && dateTimes.Contains(colName.ToString()!))
                                {
                                    try
                                    {
                                        long fileTime = BitConverter.ToInt64(bytes, 0);
                                        if (fileTime <= futureTicks && fileTime > 116444736000000000)
                                        {
                                            value = DateTime.FromFileTimeUtc(fileTime).ToString("yyyy-MM-dd HH:mm:ss");
                                        }
                                        else
                                        {
                                            value = "";
                                        }
                                    }
                                    catch
                                    {
                                        value = BitConverter.ToString(bytes).Replace("-", "");
                                    }
                                }
                                else if (unicodeField.Contains(colName.ToString()!))
                                {
                                    value = Encoding.Unicode.GetString(bytes).Replace('\0', ' ').Trim();
                                }
                                else if (uint64List.Contains(colName.ToString()!))
                                {
                                    value = BitConverter.ToUInt64(bytes);
                                }
                                else if (colName.ToString()!.Equals("System.Size", StringComparison.OrdinalIgnoreCase))
                                {
                                    long fileSize = BitConverter.ToInt64(bytes, 0);
                                    value = fileSize;
                                }
                                else if (colName.ToString()!.Equals("InvertedOnlyPids", StringComparison.OrdinalIgnoreCase))
                                {
                                    List<string> pids = [];
                                    for (int i = 0; i < bytes.Length; i += 2)
                                    {
                                        string pid = BitConverter.ToInt16(bytes, i).ToString();
                                        pids.Add(pid);
                                    }
                                    value = string.Join("|", pids);
                                }
                                else if (colName.ToString()!.Equals("System.Message.ConversationIndex", StringComparison.OrdinalIgnoreCase))
                                {
                                    List<string> indices = [];
                                    for (int i = 0; i < bytes.Length; i += 2)
                                    {
                                        string index = BitConverter.ToUInt16(bytes, i).ToString();
                                        indices.Add(index);
                                    }
                                    value = string.Join("|", indices);
                                }
                                else if (durations.Contains(colName.ToString()!))
                                {
                                    ulong time = BitConverter.ToUInt64((byte[])value, 0);
                                    TimeSpan duration = TimeSpan.FromTicks((long)time);
                                    value = string.Format("{0:D2}:{1:D2}:{2:D2}:{3:D2}", (int)duration.Days, (int)duration.Hours, duration.Minutes, duration.Seconds);
                                    var description = GetMeasurementName(time, MediaDurations);
                                    value = $"{value} - {description}";
                                }
                                else if (colName.ToString()!.Equals("System.GPS.Longitude") | colName.ToString()!.Equals("System.GPS.Latitude"))
                                {
                                    value = BitConverter.ToDouble((byte[])value!, 16);
                                }
                                else
                                {
                                    value = BitConverter.ToString(bytes).Replace("-", "");
                                }
                            }
                            if (sfgaoField.Contains(colName.ToString()!))
                            {
                                List<string> flags = GetMatchingFlags((long)value, SFGAO);
                                value = string.Join("|", flags);
                            }
                            if (colName.ToString()!.Equals("System.Activity.BackgroundColor", StringComparison.OrdinalIgnoreCase))
                            {
                                uint colorVal = (uint)(long)value;
                                Color c = Color.FromArgb((int)colorVal);
                                value = $"{value} - ARGB({c.A},{c.R},{c.G},{c.B})";
                            }
                            if (colName.ToString()!.Equals("System.FilePlaceholderStatus", StringComparison.OrdinalIgnoreCase))
                            {
                                List<string> matches = GetMatchingFlags((long)value, FilePlaceholderStates);
                                value = string.Join("|", matches);
                            }
                            if (colName.ToString()!.Equals("System.FileAttributes", StringComparison.OrdinalIgnoreCase))
                            {
                                List<string> flags = GetMatchingFlags((long)value, FileAttributes);
                                value = string.Join("|", flags);
                            }
                            if (guids.Contains(colName.ToString()!))
                            {
                                Guid guid = new((string)value);
                                value = guid.ToString("B").ToUpper();
                            }
                            if (lookups.Contains(colName.ToString()!))
                            {
                                if (LookupValues.TryGetValue(colName.ToString()!, out var lookup))
                                {
                                    if (lookup.TryGetValue((int)(long)value, out string? strValue))
                                    {
                                        value = strValue;
                                    }
                                }
                            }
                            if (colName.ToString()!.Equals("System.Capacity", StringComparison.OrdinalIgnoreCase) && value is not null)
                            {
                                string capacityName = GetMeasurementName((ulong)value, Capacity);
                                value = (string)value + capacityName;
                            }
                            if (colName.ToString()!.Equals("System.Message.Flags", StringComparison.OrdinalIgnoreCase) && value is not null)
                            {
                                List<string> matches = GetMatchingFlags((long)value, MessageFlags);
                                value = string.Join("|", matches);
                            }
                            row.Add(value!);
                        }
                        else
                        {
                            row.Add(DBNull.Value);
                        }
                    }
                    IndexResults.Add(row);
                }
            }
            catch (Exception ex)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Unable to read and correlate data:\n\n{ex.Message}", "Data Correlation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                GoButton.IsEnabled = true;
            }
        }

        private static Dictionary<int, string> BuildFullPaths(List<List<string>> paths)
        {
            var lookup = new Dictionary<int, (int Parent, string Path)>();
            foreach (var row in paths)
            {
                if (row.Count >= 3 && int.TryParse(row[0], out int scopeId) && int.TryParse(row[1], out int parentId))
                {
                    lookup[scopeId] = (parentId, row[2]);
                }
            }
            var result = new Dictionary<int, string>();
            string ResolvePathIterative(int scopeId)
            {
                if (result.TryGetValue(scopeId, out string? fullPath))
                {
                    return fullPath;
                }
                if (!lookup.ContainsKey(scopeId))
                {
                    return "";
                }
                var pathParts = new Stack<string>();
                var currentId = scopeId;
                while (lookup.TryGetValue(currentId, out var entry))
                {
                    pathParts.Push(entry.Path);
                    if (entry.Parent == 1)
                    {
                        break;
                    }
                    currentId = entry.Parent;
                }
                var sb = new StringBuilder();
                while (pathParts.Count > 0)
                {
                    sb.Append(pathParts.Pop());
                }
                fullPath = sb.ToString();
                result[scopeId] = fullPath;
                return fullPath;
            }
            foreach (var scopeId in lookup.Keys)
            {
                ResolvePathIterative(scopeId);
            }
            return result;
        }

        private async Task<bool> ExportToExcel(string filePath, Dictionary<string, List<List<object>>> dataToExport)
        {
            ExcelPackage.License.SetNonCommercialOrganization("Digital Sleuth");
            if (dataToExport == null || dataToExport.Count == 0)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show("No data to export.", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                GoButton.IsEnabled = true;
                return false;
            }

            try
            {
                await Task.Run(() => {
                    using var package = new ExcelPackage();
                    foreach (var (title, results) in dataToExport)
                    {
                        if (results is not [])
                        {
                            var worksheet = package.Workbook.Worksheets.Add(title);
                            HashSet<int> dateTimeCols = [];
                            for (int col = 0; col < results[0].Count; col++)
                            {
                                string header = results[0][col]?.ToString() ?? string.Empty;
                                int excelCol = col + 1;

                                worksheet.Cells[1, excelCol].Value = header;

                                if (dateTimes.Contains(header) || header == "Timestamp")
                                {
                                    dateTimeCols.Add(excelCol);
                                    worksheet.Column(excelCol).Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                                }
                                else
                                {
                                    worksheet.Column(excelCol).Style.Numberformat.Format = "@";
                                }
                            }
                            for (int row = 1; row < results.Count; row++)
                            {
                                var dataRow = results[row];

                                for (int col = 0; col < dataRow.Count; col++)
                                {
                                    int excelCol = col + 1;
                                    var cell = worksheet.Cells[row + 1, excelCol];

                                    if (dateTimeCols.Contains(excelCol))
                                    {
                                        if (dataRow[col] is DateTime dt)
                                        {
                                            cell.Value = dt;
                                        }
                                        else if (DateTime.TryParse(dataRow[col]?.ToString(), out var parsed))
                                        {
                                            cell.Value = parsed;
                                        }
                                        else
                                        {
                                            cell.Value = null;
                                        }
                                    }
                                    else
                                    {
                                        cell.Value = FormatValueForExcel(dataRow[col]);
                                    }
                                }
                            }
                            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                            worksheet.View.FreezePanes(2, 1);
                            worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column]
                                     .AutoFilter = true;
                        }
                    }
                    package.SaveAs(new FileInfo(filePath));
                });
                return true;
            }
            catch (Exception ex)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"An error occurred while exporting:\n\n{ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                GoButton.IsEnabled = true;
                return false;
            }
        }

        private static string FormatValueForExcel(object value)
        {
            if (value == null || value is DBNull)
            {
                return string.Empty;
            }

            if (value is byte[] bytes)
            {
                return BitConverter.ToString(bytes).Replace("-", "");
            }

            return value.ToString() ?? string.Empty;
        }

        private static List<List<object>> FilterColumns(List<List<object>> source, List<string> desiredColumns)
        {
            if (source == null || source.Count == 0)
                return [];

            var headerRow = source[0];
            var headers = headerRow.Select(h => h?.ToString() ?? string.Empty).ToList();

            var selectedIndexes = desiredColumns
                .Select(col => headers.IndexOf(col))
                .Where(index => index >= 0)
                .ToList();

            return [.. source
                .Select(row => selectedIndexes
                    .Where(i => i < row.Count)
                    .Select(i => row[i])
                    .ToList())];
        }

        private void AboutClick(object sender, RoutedEventArgs e)
        {
            App.Current.MainWindow.Activate();
            MessageBoxResult result = System.Windows.MessageBox.Show(
                $"{displayName} v{appVersion}\n" +
                "A simple Windows Search Index parser for the SQLite and ESE DB versions.\n\n" +
                $"Author: Corey Forman (digitalsleuth)\n" +
                $"Source: {githubBinaryRepo}\n\n" +
                $"Would you like to visit the repo on GitHub?",
                $"{displayName} v{appVersion}", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {
                Process.Start(new ProcessStartInfo($"{githubBinaryRepo}") { UseShellExecute = true });
            }
        }

        public class SerializedPropertyStore
        {
            public string? FormatId { get; set; }
            public Dictionary<string, object> Properties { get; set; } = [];
        }

        public static class SPSParser
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-propstore/3453fb82-0e4f-4c2c-bc04-64b4bd2c51ec
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/f122b9d7-e5cf-4484-8466-83f6fd94b3cc
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/5a2b34c4-d109-438e-9ec8-84816d8de40d
        /// CodePage Property Id? - https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/04727953-c174-4d01-80ae-b38b15d7068a
        {
            public static List<SerializedPropertyStore> Parse(byte[] blob)
            {
                string SPString = "{d5cdd505-2e9c-101b-9397-08002b2cf9ae}";
                var stores = new List<SerializedPropertyStore>();
                using var ms = new MemoryStream(blob);
                using var reader = new BinaryReader(ms);
                var firstBytes = reader.ReadInt32();
                if (reader.BaseStream.Position < reader.BaseStream.Length - 4)
                {
                    if (firstBytes == 1)
                    {
                        reader.ReadInt32();
                        List<string> dataStrings = [];
                        string data;
                        while (reader.BaseStream.Position < reader.BaseStream.Length)
                        {
                            data = ReadUnicodeString(reader, false);
                            dataStrings.Add(data);
                            if (reader.BaseStream.Position != reader.BaseStream.Length && reader.BaseStream.Length - reader.BaseStream.Position != 4)
                            {
                                reader.BaseStream.Position -= 2;
                                reader.ReadInt16();
                            }
                        }
                        var store = new SerializedPropertyStore { FormatId = "Partially Serialized Data" };
                        store.Properties["Data"] = string.Join(" ", dataStrings);
                        stores.Add(store);
                        return stores;
                    }
                    int storeSize = reader.ReadInt32();
                    long endPos = reader.BaseStream.Position - 4 + storeSize;
                    while (reader.BaseStream.Position < endPos)
                    {
                        string propertyName = "";
                        var store = new SerializedPropertyStore { FormatId = null };
                        int storageSize = reader.ReadInt32();
                        long storageEnd = reader.BaseStream.Position - 4 + storageSize;
                        string version = Encoding.ASCII.GetString(reader.ReadBytes(4));
                        if (version != "1SPS")
                        {
                            if (firstBytes == 6029404)
                            {
                                var data = Encoding.Unicode.GetString(blob).Trim('\0');
                                store.FormatId = "Raw Data";
                                store.Properties["Data"] = data;
                                stores.Add(store);
                                return stores;
                            }
                            else
                            {
                                throw new InvalidDataException("Not a valid SPS signature.");
                            }
                        }
                        string formatId = new Guid(reader.ReadBytes(16)).ToString("B");
                        store.FormatId = formatId;
                        uint spValueSize = reader.ReadUInt32();
                        if (formatId == SPString)
                        {
                            int spNameSize = reader.ReadInt32();
                            _ = reader.ReadBytes(1);
                            string spNameVar = Encoding.Unicode.GetString(reader.ReadBytes(spNameSize));
                            store.Properties["Name"] = spNameVar;
                            byte[] spValueVar = reader.ReadBytes((int)spValueSize);
                            string spValueString = BitConverter.ToString(spValueVar).Replace("-", "").TrimEnd('\0');
                            store.Properties["Value"] = spValueString;
                        }
                        else
                        {
                            uint spId = reader.ReadUInt32();
                            reader.ReadBytes(1);
                            foreach (var entry in PropertyKeys)
                            {
                                if (entry.Value.guid == formatId && entry.Value.propertyId == spId)
                                {
                                    propertyName = entry.Key;
                                    store.Properties["Name"] = propertyName;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            store.Properties["PropertyId"] = spId.ToString();
                            reader.ReadUInt32();
                            uint spSize = reader.ReadUInt32();
                            reader.ReadUInt16();
                            ushort vt = reader.ReadUInt16();
                            reader.ReadUInt16();
                            object? spValue = ReadPropertyValue(reader, vt);
                            if (propertyName == "Activity_ContentVisualPropertiesHash")
                            {
                                byte[] byteValue = BitConverter.GetBytes((ulong)spValue!);
                                spValue = BitConverter.ToString(byteValue).Replace("-", "");
                            }
                            if (propertyName == "SFGAOFlags")
                            {
                                List<string> flags = GetMatchingFlags((long)(uint)spValue!, SFGAO);
                                spValue = string.Join("|", flags);
                            }
                            store.Properties["Value"] = spValue!;
                            store.Properties["VT_ID"] = VariantType[vt];
                        }
                        stores.Add(store);
                        reader.BaseStream.Seek(storageEnd, SeekOrigin.Begin);
                    }
                }
                return stores;
            }

            private static readonly JsonSerializerOptions writeOptions = new()
            {
                WriteIndented = true,
            };

            public static string ParseToJson(byte[] blob)
            {
                var stores = Parse(blob);
                return JsonSerializer.Serialize(stores, writeOptions);
            }

            private static object? ReadPropertyValue(BinaryReader reader, int vt)
            {
                /// Not all VT's are set here because they are not all used in the Windows Search Index.
                return vt switch
                {
                    0x00 => null, // VT_EMPTY
                    0x01 => null, // VT_NULL
                    0x02 => reader.ReadInt16(), // VT_I2
                    0x03 => reader.ReadInt32(), // VT_I4
                    0x04 => reader.ReadSingle(), // VT_R4
                    0x05 => reader.ReadDouble(), // VT_R8
                    0x06 => reader.ReadInt64(), // VT_CY
                    0x07 => reader.ReadDouble(), // VT_DATE
                    0x08 => ReadString(reader), // VT_BSTR
                    0x0A => reader.ReadUInt32(), // VT_ERROR
                    0x0B => reader.ReadInt32() != 0, // VT_BOOL
                    0x0E => ReadDecimal(reader), // VT_DECIMAL
                    0x10 => reader.ReadSByte(), // VT_I1
                    0x11 => reader.ReadUInt32(), // VT_UI1
                    0x12 => reader.ReadUInt16(), // VT_UI2
                    0x13 => reader.ReadUInt32(), // VT_UI4
                    0x14 => reader.ReadInt64(), // VT_I8
                    0x15 => reader.ReadUInt64(), // VT_UI8
                    0x16 => reader.ReadInt32(), // VT_INT
                    0x17 => reader.ReadUInt32(), // VT_UINT
                    0x1E => ReadString(reader), // VT_LPSTR
                    0x1F => ReadUnicodeString(reader, true), // VT_LPWSTR
                    0x40 => ReadFileTime(reader), // VT_FILETIME
                    0x41 => ReadBlob(reader), // VT_BLOB
                    0x42 => ReadUnicodeString(reader, true), // VT_STREAM
                    0x43 => ReadUnicodeString(reader, true), // VT_STORAGE
                    0x44 => ReadUnicodeString(reader, true), // VT_STREAMED_Object
                    0x45 => ReadUnicodeString(reader, true), // VT_STORED_Object
                    0x46 => ReadBlob(reader), // VT_BLOB_Object
                    0x47 => ReadClipboardData(reader), // VT_CF
                    0x48 => ReadGuid(reader), // VT_CLSID
                    0x49 => ReadVersionedStream(reader), // VT_VERSIONED_STREAM
                    0x101F => ReadVectorLpwstr(reader), // VT_VECTOR | VT_LPWSTR
                    _ => $"[Unsupported VT {vt:X}"
                };
            }

            private static DateTime ReadFileTime(BinaryReader reader)
            {
                return DateTime.FromFileTimeUtc(reader.ReadInt64());
            }

            private static Guid ReadGuid(BinaryReader reader)
            {
                return new Guid(reader.ReadBytes(16));
            }

            private static string ReadVectorLpwstr(BinaryReader reader)
            {
                // 2-byte buffers of null at the end of each property
                // Some entries don't align properly to the 2-byte buffer at the end, 
                // so an assessment has to be made if the next bytes are 0's and
                // re-align if necessary.
                var result = new List<string>();
                uint count = reader.ReadUInt32();
                long streamLength = reader.BaseStream.Length;
                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        uint charCount = reader.ReadUInt32();
                        byte[] strBytes = reader.ReadBytes((int)charCount * 2);
                        string value = Encoding.Unicode.GetString(strBytes).TrimEnd('\0');
                        result.Add(value);
                        long pos = reader.BaseStream.Position;
                        if (pos != streamLength)
                        {
                            long buffer = BitConverter.ToInt16(reader.ReadBytes(2));
                            if (buffer != 0)
                            {
                                reader.BaseStream.Seek(pos, SeekOrigin.Begin);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"An error has occurred while attempting to read a VectorLpwstr:\n\n{ex.Message}");
                    }
                }
                string stringVal = string.Join(" ", result);
                return stringVal;
            }

            private static string ReadVersionedStream(BinaryReader reader)
            {
                var guid = new Guid(reader.ReadBytes(16));
                uint charCount = reader.ReadUInt32();
                byte[] strBytes = reader.ReadBytes((int)charCount * 2);
                string streamName = Encoding.Unicode.GetString(strBytes);
                long pos = reader.BaseStream.Position;
                long aligned = (pos + 3) & ~3;
                reader.BaseStream.Seek(aligned, SeekOrigin.Begin);
                return $"{guid} - {streamName}";

            }

            private static string ReadClipboardData(BinaryReader reader)
            {
                int length = reader.ReadInt32();
                var bytes = reader.ReadBytes(length);
                return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            }

            private static string ReadUnicodeString(BinaryReader reader, bool dbl)
            {
                int length = reader.ReadInt32();
                if (dbl)
                {
                    length *= 2;
                }
                var bytes = reader.ReadBytes(length);
                return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            }

            private static string ReadString(BinaryReader reader)
            {
                int length = reader.ReadInt32();
                var bytes = reader.ReadBytes(length);
                return Encoding.ASCII.GetString(bytes).TrimEnd('\0');
            }

            private static string ReadBlob(BinaryReader reader)
            {
                int length = reader.ReadInt32();
                var bytes = reader.ReadBytes((int)length);
                return BitConverter.ToString(bytes).Replace("-", "").TrimEnd('\0');
            }

            public static decimal ReadDecimal(BinaryReader reader)
            {
                byte[] bytes = reader.ReadBytes(16);
                if (bytes.Length < 16)
                    throw new ArgumentException($"Need at least 16 bytes, only {bytes.Length} available.");
                ushort _ = BitConverter.ToUInt16(bytes, 0);
                byte scale = bytes[2];
                byte sign = bytes[3];
                uint hi32 = BitConverter.ToUInt32(bytes, 4);
                ulong lo64 = BitConverter.ToUInt64(bytes, 8);

                int lo = (int)(lo64 & 0xFFFFFFFF);
                int mid = (int)((lo64 >> 32) & 0xFFFFFFFF);
                int hi = (int)hi32;

                bool isNegative = (sign & 0x80) != 0;

                return new decimal(lo, mid, hi, isNegative, scale);
            }
        }
    }

    public class LibEsedbReader : IDisposable
    {
        private IntPtr _fileHandle = IntPtr.Zero;
        private bool _isOpen = false;

        #region P/Invoke Declarations

        // File operations
        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int libesedb_file_initialize(out IntPtr file, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int libesedb_file_free(ref IntPtr file, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        private static extern int libesedb_file_open_wide(IntPtr file, string filename, int access_flags, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_file_close(IntPtr file, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_file_get_number_of_tables(IntPtr file, out int number_of_tables, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_file_get_table(IntPtr file, int table_entry, out IntPtr table, out IntPtr error);

        // Table operations
        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_free(ref IntPtr table, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_utf8_name_size(IntPtr table, out UIntPtr utf8_name_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_utf8_name(IntPtr table, byte[] utf8_name, UIntPtr utf8_name_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_number_of_columns(IntPtr table, out int number_of_columns, int flags, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_column(IntPtr table, int column_entry, out IntPtr column, int flags, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_number_of_records(IntPtr table, out int number_of_records, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_record(IntPtr table, int record_entry, out IntPtr record, out IntPtr error);

        // Column operations
        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_column_free(ref IntPtr column, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_column_get_utf8_name_size(IntPtr column, out UIntPtr utf8_name_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_column_get_utf8_name(IntPtr column, byte[] utf8_name, UIntPtr utf8_name_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_column_get_type(IntPtr column, out uint column_type, out IntPtr error);

        // Record operations
        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_free(ref IntPtr record, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_is_long_value(IntPtr record, int value_entry, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_is_multi_value(IntPtr record, int value_entry, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_number_of_values(IntPtr record, out int number_of_values, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value(IntPtr record, int value_entry, out IntPtr record_value, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_column_type(IntPtr record, int value_entry, out uint column_type, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_utf8_value_size(IntPtr record, int value_entry, out UIntPtr utf8_value_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_utf8_string_size(IntPtr record, int value_entry, out UIntPtr utf8_string_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_utf8_string(IntPtr record, int value_entry, byte[] utf8_value, UIntPtr utf8_value_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_data_size(IntPtr record, int value_entry, out UIntPtr value_data_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_data(IntPtr record, int value_entry, byte[] value_data, UIntPtr value_data_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_data_flags(IntPtr record, int value_entry, out byte value_data_flags, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_32bit(IntPtr record, int value_entry, out uint value_32bit, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_64bit(IntPtr record, int value_entry, out ulong value_64bit, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_data(IntPtr multi_value, int multi_value_index, byte[] value_data, nuint value_data_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_data_size(IntPtr multi_value, int multi_value_index, out nuint value_data_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_number_of_values(IntPtr multi_value, out int number_of_values, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_multi_value(IntPtr record, int value_entry, out IntPtr multi_value, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_free(ref IntPtr multi_value, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_binary_data_size(IntPtr multi_value, int multi_value_index, out nuint value_data_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_binary_data(IntPtr multi_value, int multi_value_index, byte[] value_data, nuint value_data_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_utf8_string_size(IntPtr multi_value, int multi_value_index, out nuint utf8_string_size, out IntPtr error);

        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_utf8_string(IntPtr multi_value, int multi_value_index, byte[] utf8_string, nuint utf8_string_size, out IntPtr error);
        // Error handling
        [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_error_free(ref IntPtr error);
        // Compression - not yet available
        // [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        // private static extern int libesedb_compression_7bit_decompress_get_size(byte[] compressed_data, UIntPtr compressed_data_size, out UIntPtr uncompressed_data_size, IntPtr error);

        // [DllImport("libesedb.dll", CallingConvention = CallingConvention.Cdecl)]
        // private static extern int libesedb_compression_7bit_decompress(byte[] compressed_data, UIntPtr compressed_data_size, byte[] uncompressed_data, UIntPtr uncompressed_data_size, IntPtr error);

        #endregion

        public void Open(string databasePath)
        {
            if (_isOpen)
                return;

            int result = libesedb_file_initialize(out _fileHandle, out nint error);
            if (result != 1)
            {
                throw new Exception("Failed to initialize libesedb file");
            }
            result = libesedb_file_open_wide(_fileHandle, databasePath, 0x01, out error);
            if (result != 1)
            {
                _ = libesedb_file_free(ref _fileHandle, out error);
                throw new Exception($"Failed to open database: {databasePath}");
            }

            _isOpen = true;
        }

        public List<string> GetTableNames()
        {
            EnsureOpen();
            var tableNames = new List<string>();
            int result = libesedb_file_get_number_of_tables(_fileHandle, out int numberOfTables, out nint error);
            if (result != 1)
            {
                throw new Exception("Failed to get number of tables");
            }
            for (int i = 0; i < numberOfTables; i++)
            {
                result = libesedb_file_get_table(_fileHandle, i, out nint tableHandle, out error);

                if (result == 1 && tableHandle != IntPtr.Zero)
                {
                    string? tableName = GetTableName(tableHandle);
                    if (!string.IsNullOrEmpty(tableName))
                    {
                        tableNames.Add(tableName);
                    }
                    _ = libesedb_table_free(ref tableHandle, out error);
                }
            }

            return tableNames;
        }

        public Dictionary<string, uint> GetColumnInfo(string tableName)
        {
            EnsureOpen();
            var columns = new Dictionary<string, uint>();
            IntPtr tableHandle = GetTableHandle(tableName);

            if (tableHandle == IntPtr.Zero)
            {
                return columns;
            }
            try
            {
                int result = libesedb_table_get_number_of_columns(tableHandle, out int numberOfColumns, 0, out nint error);
                if (result == 1)
                {
                    for (int i = 0; i < numberOfColumns; i++)
                    {
                        result = libesedb_table_get_column(tableHandle, i, out nint columnHandle, 0, out error);

                        if (result == 1 && columnHandle != IntPtr.Zero)
                        {
                            string? columnName = GetColumnName(columnHandle);
                            _ = libesedb_column_get_type(columnHandle, out uint columnType, out error);
                            if (!string.IsNullOrEmpty(columnName))
                            {
                                columns[columnName] = columnType;
                            }
                            _ = libesedb_column_free(ref columnHandle, out error);
                        }
                    }
                }
            }
            finally
            {
                _ = libesedb_table_free(ref tableHandle, out nint error);
            }

            return columns;
        }

        public Dictionary<string, List<object>> ReadAllData()
        {
            EnsureOpen();
            var allData = new Dictionary<string, List<object>>();
            var tableNames = GetTableNames();

            foreach (var tableName in tableNames)
            {
                try
                {
                    var tableData = ReadTableData(tableName);
                    allData[tableName] = tableData;
                }
                catch
                {
                    allData[tableName] = [];
                }
            }
            return allData;
        }

        public List<object> ReadTableData(string tableName)
        {
            EnsureOpen();
            var rows = new List<object>();
            IntPtr tableHandle = GetTableHandle(tableName);

            if (tableHandle == IntPtr.Zero)
            {
                return rows;
            }
            try
            {
                var columnInfo = new List<string>();
                int result = libesedb_table_get_number_of_columns(tableHandle, out int numberOfColumns, 0, out nint error);

                if (result == 1)
                {
                    for (int i = 0; i < numberOfColumns; i++)
                    {
                        result = libesedb_table_get_column(tableHandle, i, out nint columnHandle, 0, out error);

                        if (result == 1 && columnHandle != IntPtr.Zero)
                        {
                            string? columnName = GetColumnName(columnHandle);
                            columnInfo.Add(columnName ?? $"Column{i}");
                            _ = libesedb_column_free(ref columnHandle, out error);
                        }
                    }
                }
                result = libesedb_table_get_number_of_records(tableHandle, out int numberOfRecords, out error);
                rows.Add(columnInfo);
                if (result == 1)
                {
                    for (int i = 0; i < numberOfRecords; i++)
                    {
                        result = libesedb_table_get_record(tableHandle, i, out nint recordHandle, out error);

                        if (result == 1 && recordHandle != IntPtr.Zero)
                        {
                            var row = ReadRecord(recordHandle, columnInfo);
                            if (tableName == "SystemIndex_GthrPth")
                            {
                                List<string> strings = [.. row.Select(o => o?.ToString() ?? string.Empty)];
                                rows.Add(strings);
                            }
                            else
                            {
                                rows.Add(row);
                            }
                            _ = libesedb_record_free(ref recordHandle, out error);
                        }
                    }
                }
            }
            finally
            {
                _ = libesedb_table_free(ref tableHandle, out nint error);
            }
            return rows;
        }

        #region Helper Methods
        private IntPtr GetTableHandle(string tableName)
        {
            int result = libesedb_file_get_number_of_tables(_fileHandle, out int numberOfTables, out nint error);

            if (result != 1)
                return IntPtr.Zero;
            for (int i = 0; i < numberOfTables; i++)
            {
                result = libesedb_file_get_table(_fileHandle, i, out nint tableHandle, out error);

                if (result == 1 && tableHandle != IntPtr.Zero)
                {
                    string? currentName = GetTableName(tableHandle);
                    if (currentName == tableName)
                    {
                        return tableHandle;
                    }

                    _ = libesedb_table_free(ref tableHandle, out error);
                }
            }
            return IntPtr.Zero;
        }

        private static string? GetTableName(IntPtr tableHandle)
        {
            int result = libesedb_table_get_utf8_name_size(tableHandle, out nuint nameSize, out nint error);
            if (result != 1 || nameSize.ToUInt64() == 0)
                return null;

            byte[] nameBuffer = new byte[nameSize.ToUInt64()];
            result = libesedb_table_get_utf8_name(tableHandle, nameBuffer, nameSize, out error);
            if (result != 1)
                return null;
            return Encoding.UTF8.GetString(nameBuffer).TrimEnd('\0');
        }

        private static string? GetColumnName(IntPtr columnHandle)
        {
            int result = libesedb_column_get_utf8_name_size(columnHandle, out nuint nameSize, out nint error);
            if (result != 1 || nameSize.ToUInt64() == 0)
                return null;

            byte[] nameBuffer = new byte[nameSize.ToUInt64()];
            result = libesedb_column_get_utf8_name(columnHandle, nameBuffer, nameSize, out error);
            if (result != 1)
                return null;
            return Encoding.UTF8.GetString(nameBuffer).TrimEnd('\0');
        }

        private static List<object> ReadRecord(IntPtr recordHandle, List<string> columnNames)
        {
            var row = new List<object>();
            int result = libesedb_record_get_number_of_values(recordHandle, out int numberOfValues, out nint error);
            if (result != 1)
                return row;

            for (int i = 0; i < numberOfValues && i < columnNames.Count; i++)
            {
                object? value = GetRecordValue(recordHandle, i);
                row.Add(value!);
            }
            return row;
        }

        private static object? GetRecordValue(IntPtr recordHandle, int valueIndex)
        {
            int result = libesedb_record_get_column_type(recordHandle, valueIndex, out _, out nint error);
            if (result != 1)
                return null;

            result = libesedb_record_get_value_utf8_string_size(recordHandle, valueIndex, out nuint valueSize, out error);
            if (result == 1 && valueSize.ToUInt64() > 0)
            {
                byte[] valueBuffer = new byte[valueSize.ToUInt64()];
                result = libesedb_record_get_value_utf8_string(recordHandle, valueIndex, valueBuffer, valueSize, out error);
                if (result == 1)
                {
                    return Encoding.UTF8.GetString(valueBuffer).TrimEnd('\0');
                }
            }

            result = libesedb_record_get_value_32bit(recordHandle, valueIndex, out uint value32, out error);
            if (result == 1)
            {
                return value32;
            }

            result = libesedb_record_get_value_64bit(recordHandle, valueIndex, out ulong value64, out error);
            if (result == 1)
            {
                return value64;
            }

            result = libesedb_record_get_value_data_size(recordHandle, valueIndex, out valueSize, out error);
            if (result == 1 && valueSize.ToUInt64() > 0)
            {
                byte[] dataBuffer = new byte[valueSize.ToUInt64()];
                result = libesedb_record_get_value_data(recordHandle, valueIndex, dataBuffer, valueSize, out error);
                if (result == 1)
                {
                    return dataBuffer;
                }
            }
            return null;
        }

        private void EnsureOpen()
        {
            if (!_isOpen)
                throw new InvalidOperationException("Database is not open. Call Open() first.");
        }

        #endregion

        public void Dispose()
        {
            if (_isOpen && _fileHandle != IntPtr.Zero)
            {
                _ = libesedb_file_close(_fileHandle, out nint error);
                _ = libesedb_file_free(ref _fileHandle, out error);
                _isOpen = false;
            }
        }
    }
}
