ENTERPRISE FILE MANAGER v3.3
================================

EFM.exe is a portable Windows application. It does not require installation,
PowerShell, Python, administrator rights, or an internet connection.

USAGE
1. Double-click EFM.exe.
2. Select the folder containing the documents you want to manage.
3. Select a document and use Open File to view it in its Windows-associated application.
4. Choose Update Details to save changes.

Use Refresh List to rescan the selected folder. The gear icon in the top-right
opens Settings, where Manage Groups and Manage Status control the available
definitions and their supporting descriptions.
Use Manage List to open Document Management. Find Possible Duplicates compares
SHA-256 hashes and supporting content, filename, size, and metadata signals.
Suggestions are read-only until the user reviews a pair, chooses the file to
keep, and confirms the move. Confidence is guidance rather than proof.
Analysis fingerprints each file once and scores focused candidate groups rather
than comparing every file with every other file. It runs in the background with
stage-based progress and can be stopped using Cancel Analysis.
When removing a Group or Status definition that is currently assigned, EFM
warns how many documents are affected and lists the affected document names.

The Document Title field is saved with each document's metadata and displayed
in the Document Title column. The search box matches both file names and titles.
Use the Status filter above the document list to show all documents or only
Current, Legacy (Still Relevant), or Historic (Archived) documents.
Selecting a different document opens its details on the Metadata tab. Saving
or continuing to edit the same document keeps the currently active details tab.
Each document-details tab provides an Update Details button for saving changes.
The Network tab shows the selected document and only the documents listed in
its Associations tab. Click an associated document node to open that document's
details, then select its Network tab to continue exploring direct associations.
Associations are reciprocal: linking document A to document B automatically
links document B to document A. Removing the link removes both directions.
Association and version-history lists display both the document name and title.
Click any column heading in these lists to sort it; click again to reverse the
order. An arrow on the heading indicates the active sort direction.
The Groups tab assigns a document to one or more predefined logical groups and
displays each assigned group's short description in a separate column.
New group definitions can only be created through Settings > Manage Groups.
The Metadata tab assigns Status from the list defined in Settings > Manage Status.
Hover over assigned groups and status values to view their descriptions.

Use the Associations and Version History tabs to link related documents. Links
open the selected document inside EFM; use Back to return to the prior record.
Document issue dates are stored alongside the remaining document details.
The Add Documents windows support Ctrl/Shift multi-selection so several
associations or version-history links can be added in one operation.
The Attachments tab links one or more files from the selected folder to a
document. Attached files remain in the folder and can be opened from the tab,
but are hidden from the Documents list and reconciliation review until removed
from the attachment list.

The application stores document metadata in a file named filedata.efm inside
the selected document folder. Existing filedata.efm files created by the
PowerShell version remain compatible.
Folder-specific group and status definitions are stored in settings.efs in the
same folder. Existing settings.efm and groups.efm files are migrated automatically.
Duplicate-review audit records are stored separately in duplicates.efs. EFM
creates this versioned audit file when a document folder is first selected and
excludes it from the Documents list.
Confirmed duplicate files are moved into the selected folder's duplicates
subfolder. Metadata references are redirected to the retained file, and the
Resolved Duplicates view can restore an audited move to its original location.

When a folder is opened or refreshed, EFM checks for files without metadata
and metadata records without a corresponding file. Any discrepancies are shown
for review. A file without metadata can be matched to a proposed or manually
selected orphaned metadata record after explicit confirmation. Other records
remain unchanged.

PORTABILITY
You can copy EFM.exe to another Windows 10/11 x64 computer and run it directly.
Windows may display a SmartScreen warning because the executable is not
digitally code-signed. This does not indicate that installation is required.

DATA SAFETY
EFM does not move, rename, or delete managed documents during ordinary editing.
Document Management moves a file only after explicit duplicate confirmation;
the move and affected metadata are recorded in duplicates.efs for restoration.
