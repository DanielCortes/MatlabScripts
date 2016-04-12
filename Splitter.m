% Demo macro to write numerical arrays and cell arrays
% to two different worksheets in an Excel workbook file.
% Uses xlswrite1, available from the File Exchange
% http://www.mathworks.com/matlabcentral/fileexchange/10465-xlswrite1
function Splitter
clc;
close all;
clear all;
fullFileName = GetXLFileName();
if isempty(fullFileName)
	% User clicked Cancel.
	return;
end
Excel = actxserver('Excel.Application');   

% Prepare proper filename extension.
% Get the Excel version because if it's version 11 (Excel 2003) the file extension should be .xls,
% but if it's 12.0 (Excel 2007) then we'll need to use an extension of .xlsx to avoid nag messages.
excelVersion = str2double(Excel.Version);
if excelVersion < 12
	excelExtension = '.xls';
else
	excelExtension = '.xlsx';
end

% Determine the proper format to save the files in.  It depends on the extension (Excel version).
switch excelExtension
	case '.xls' %xlExcel8 or xlWorkbookNormal
	   xlFormat = -4143;
	case '.xlsb' %xlExcel12
	   xlFormat = 50;
	case '.xlsx' %xlOpenXMLWorkbook
	   xlFormat = 51;
	case '.xlsm' %xlOpenXMLWorkbookMacroEnabled 
	   xlFormat = 52;
	otherwise
	   xlFormat = -4143;
end

if ~exist(fullFileName, 'file')
	message = sprintf('I am going to create Excel workbook:\n\n%s\n\nClick OK to continue.\nClick Exit to exit this function', fullFileName);
	button = questdlg(message, 'Creating new workbook', 'OK', 'Exit', 'OK');
	drawnow;	% Refresh screen to get rid of dialog box remnants.
	if strcmpi(button, 'Exit')
		return;
	end
	% Add a new workbook.
	ExcelWorkbook = Excel.workbooks.Add;
	% Save this workbook we just created.
	ExcelWorkbook.SaveAs(fullFileName, xlFormat);
	ExcelWorkbook.Close(false);
else
	% Delete the existing file.
	delete(fullFileName);
end

% Open up the workbook named in the variable fullFileName.
invoke(Excel.Workbooks,'Open', fullFileName);
Excel.visible = true;

% Create some sample data.

load('C:\Users\Dan\Desktop\GD5_001mid\TrackingPackage\tracks\Channel_1_tracking_result_mat.mat')
SizeVec = size(trackedFeatureInfo);
Size = SizeVec(1);
po = 0;
n = 0;
for i = 1:Size
    n = 9 * po + 1;
    Mat = trackedFeatureInfo(i,:);
    Matt = reshape(Mat, 8, []);
    my_cell = sprintf('A%s',num2str(n));
    xlswrite1( fullFileName, Matt, 1, my_cell);
    po = po + 1;
    prog = (po / Size) * 100;
    fprintf('Progress: %5.2f\n', prog)
end

% Delete all empty sheets in the active workbook.
DeleteEmptyExcelSheets(Excel);

% Then run the following code to close the activex server:
invoke(Excel.ActiveWorkbook,'Save');
Excel.Quit;
Excel.delete;
clear Excel;
message = sprintf('Done!\nThis Excel workbook has been created:\n%s', fullFileName);
msgbox(message);
% End of main function: ExcelDemo.m -----------------------------

%--------------------------------------------------------------------
% Gets the name of the workbook from the user.
function fullExcelFileName = GetXLFileName()
	fullExcelFileName = [];  % Default.
	% Ask user for a filename.
	FilterSpec = {'*.xls', 'Excel workbooks (*.xls)'; '*.*', 'All Files (*.*)'};
	DialogTitle = 'Save workbook file name';
	% Get the default filename.  Make sure it's in the folder where this m-file lives.
	% (If they run this file but the cd is another folder then pwd will show that folder, not this one.
	thisFile = mfilename('fullpath');
	[thisFolder, baseFileName, ext] = fileparts(thisFile);
	DefaultName = sprintf('%s/%s.xls', thisFolder, baseFileName);
	[fileName, specifiedFolder] = uiputfile(FilterSpec, DialogTitle, DefaultName);
	if fileName == 0
		% User clicked Cancel.
		return;
	end
	% Parse what they actually specified.
	[folder, baseFileName, ext] = fileparts(fileName);
	% Create the full filename, making sure it has a xls filename.
	fullExcelFileName = fullfile(specifiedFolder, [baseFileName '.xls']);

% --------------------------------------------------------------------
% DeleteEmptyExcelSheets: deletes all empty sheets in the active workbook.
% This function looped through all sheets and deletes those sheets that are
% empty. Can be used to clean a newly created xls-file after all results
% have been saved in it.
function DeleteEmptyExcelSheets(excelObject)
% 	excelObject = actxserver('Excel.Application');
% 	excelWorkbook = excelObject.workbooks.Open(fileName);
	worksheets = excelObject.sheets;
	sheetIdx = 1;
	sheetIdx2 = 1;
	numSheets = worksheets.Count;
	% Prevent beeps from sounding if we try to delete a non-empty worksheet.
	excelObject.EnableSound = false;

	% Loop over all sheets
	while sheetIdx2 <= numSheets
		% Saves the current number of sheets in the workbook
		temp = worksheets.count;
		% Check whether the current worksheet is the last one. As there always
		% need to be at least one worksheet in an xls-file the last sheet must
		% not be deleted.
		if or(sheetIdx>1,numSheets-sheetIdx2>0)
			% worksheets.Item(sheetIdx).UsedRange.Count is the number of used cells.
			% This will be 1 for an empty sheet.  It may also be one for certain other
			% cases but in those cases, it will beep and not actually delete the sheet.
			if worksheets.Item(sheetIdx).UsedRange.Count == 1
				worksheets.Item(sheetIdx).Delete;
			end
		end
		% Check whether the number of sheets has changed. If this is not the
		% case the counter "sheetIdx" is increased by one.
		if temp == worksheets.count;
			sheetIdx = sheetIdx + 1;
		end
		sheetIdx2 = sheetIdx2 + 1; % prevent endless loop...
	end
	excelObject.EnableSound = true;
	return;

