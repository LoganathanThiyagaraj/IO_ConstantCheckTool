%*****************************************************
% Tool     : IO_ConstantCheckTool
% Developer: Loganathan Thiyagaraj
% Owner    : Noopur Dosi
%*****************************************************

function varargout = IO_ConstantCheckTool(varargin)
% IO_CONSTANTCHECKTOOL MATLAB code for IO_ConstantCheckTool.fig
%      IO_CONSTANTCHECKTOOL, by itself, creates a new IO_CONSTANTCHECKTOOL or raises the existing
%      singleton*.
%
%      H = IO_CONSTANTCHECKTOOL returns the handle to a new IO_CONSTANTCHECKTOOL or the handle to
%      the existing singleton*.
%
%      IO_CONSTANTCHECKTOOL('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in IO_CONSTANTCHECKTOOL.M with the given input arguments.
%
%      IO_CONSTANTCHECKTOOL('Property','Value',...) creates a new IO_CONSTANTCHECKTOOL or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before IO_ConstantCheckTool_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to IO_ConstantCheckTool_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help IO_ConstantCheckTool

% Last Modified by GUIDE v2.5 16-Apr-2020 15:09:13

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @IO_ConstantCheckTool_OpeningFcn, ...
                   'gui_OutputFcn',  @IO_ConstantCheckTool_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before IO_ConstantCheckTool is made visible.
function IO_ConstantCheckTool_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to IO_ConstantCheckTool (see VARARGIN)

% Choose default command line output for IO_ConstantCheckTool
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes IO_ConstantCheckTool wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = IO_ConstantCheckTool_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on selection change in SWClistbox.
function SWClistbox_Callback(hObject, eventdata, handles)
% hObject    handle to SWClistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns SWClistbox contents as cell array
%        contents{get(hObject,'Value')} returns selected item from SWClistbox
global SWC
swc=get(handles.SWClistbox,'String');%Gets the list of contents as array  in list box
idx=get(handles.SWClistbox,'Value');%Gets index of SWC component in list box
SWC=swc{idx}; % Storing selected SWC in variable

% --- Executes during object creation, after setting all properties.
function SWClistbox_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SWClistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in ROM_partnumberlistbox.
function ROM_partnumberlistbox_Callback(hObject, eventdata, handles)
% hObject    handle to ROM_partnumberlistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns ROM_partnumberlistbox contents as cell array
%        contents{get(hObject,'Value')} returns selected item from ROM_partnumberlistbox
global ROMPartnumber
partnumber=get(handles.ROM_partnumberlistbox,'String');%Gets the list of contents as array  in list box
idx=get(handles.ROM_partnumberlistbox,'Value');%Gets index of ROM part number in list box
ROMPartnumber=partnumber{idx};% Storing selected ROM part number in variable

% --- Executes during object creation, after setting all properties.
function ROM_partnumberlistbox_CreateFcn(hObject, eventdata, handles)
% hObject    handle to ROM_partnumberlistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in FixedROM_partnumlistbox.
function FixedROM_partnumlistbox_Callback(hObject, eventdata, handles)
% hObject    handle to FixedROM_partnumlistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns FixedROM_partnumlistbox contents as cell array
%        contents{get(hObject,'Value')} returns selected item from FixedROM_partnumlistbox
global FixedROMPartnumber
partnumber=get(handles.FixedROM_partnumlistbox,'String');%Gets the list of contents as array  in list box
idx=get(handles.FixedROM_partnumlistbox,'Value');%Gets index of Fixed ROM part number in list box
FixedROMPartnumber=partnumber{idx};% Storing selected Fixed ROM part number in variable

% --- Executes during object creation, after setting all properties.
function FixedROM_partnumlistbox_CreateFcn(hObject, eventdata, handles)
% hObject    handle to FixedROM_partnumlistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in BrowseCSVfile.
function BrowseCSVfile_Callback(hObject, eventdata, handles)
% hObject    handle to BrowseCSVfile (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global present ConstantListfile selBasepath file1;

clc
selBasepath = uigetdir(path,'Select CSV file folder');%Saving CSV file folder path
addpath(genpath(selBasepath))%Adds CSV fiel folder path to the current folder
present=pwd;
[file1,path1]=uigetfile('*.csv','Select  SWC ConstantListfile');%Selects CSV file
file=strcat('\',file1);
file=strcat(path1,file);
ConstantListfile=file;
set(handles.csvtextedit,'String',file1);%Displays CSV file name in gui
set(handles.BrowseCSVfile,'Enable','off')%Disable browse button

% --- Executes on button press in BrowseDBfile.
function BrowseDBfile_Callback(hObject, eventdata, handles)
% hObject    handle to BrowseDBfile (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global ConstantDBfile file2;
[file2,path1]=uigetfile('*.xlsx','Select ConstantDB file');%Selects ConstDB file
file=strcat('\',file2);
file=strcat(path1,file);
ConstantDBfile=file;
set(handles.dbtextedit,'String',file2);%Displays DB file name in gui
set(handles.BrowseDBfile,'Enable','off')%Disable browse button

% --- Executes on button press in LoadSWC_partnumbers.
function LoadSWC_partnumbers_Callback(hObject, eventdata, handles)
% hObject    handle to LoadSWC_partnumbers (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global ConstantDBfile ROMTable FixedromTable selBasepath;
exl = actxserver('excel.application');
exlWkbk = exl.Workbooks;
exlFile = exlWkbk.Open(ConstantDBfile);%Opens ConstantDBfile
exlFile.Activate
ConstantDBsht=exlFile.Sheets.Item('ROM');%Selects ROM sheet
range1 = get(ConstantDBsht,'Range','A8:UA20');%Selects range to extract
range1.AutoFilter%Removes Autofilter
range1.AutoFilter
clc
ROMTable = readtable(ConstantDBfile,'Sheet','ROM','Range','RD8:UA500', 'ReadVariableNames', true);%Data from ROM sheet is stored in matlab table
ConstantDBsht=exlFile.Sheets.Item('Fixed_ROM');
range1 = get(ConstantDBsht,'Range','A:CC');%Selects range in Fixed_ROM sheet
range1.AutoFilter%Removes Autofilter in Fixed_ROM sheets
range1.AutoFilter
clc
FixedromTable = readtable(ConstantDBfile,'Sheet','Fixed_ROM','Range','A2:CC500', 'ReadVariableNames', true);

exlFile.Save();

exlFile.Close();
 
exl.Quit;
exl.delete;

 dinfo = dir([selBasepath '\*.csv']);%Stores  CSV files in variable
for k=1:numel(dinfo)
 data{k}=dinfo(k).name;
 swclist{k}=regexprep(data{k},'__.*','');%Populates SWC component names
end
 
%swclist
idx3 = find(strcmp(ROMTable.Properties.VariableNames,'L21BPRC'));
for i=idx3:length(ROMTable.Properties.VariableNames)
    if strfind(ROMTable.Properties.VariableNames{1,i},'Var')
        break;
    end
end    
idx4=i-1;
a=1;
for i=idx3:idx4
    partnumber{a}=ROMTable.Properties.VariableNames{i};%Populates ROM partnumber lists
    a=a+1;
end 
%partnumber
idx3 = find(strcmp(FixedromTable.Properties.VariableNames,'L21BPRC'));
for i=idx3:length(FixedromTable.Properties.VariableNames)
    if strfind(FixedromTable.Properties.VariableNames{1,i},'Var')
        break;
    end
end
idx4=i-1;
a=1;
for i=idx3:idx4
    fixedpartnumber{a}=FixedromTable.Properties.VariableNames{i};%Populates Fixed ROM partnumber lists
    a=a+1;
end 
set(handles.SWClistbox,'String',swclist);%Displays SWC list in listbox
set(handles.ROM_partnumberlistbox,'String',partnumber);%Displays ROM partnumber list in listbox
set(handles.FixedROM_partnumlistbox,'String',partnumber);%Displays Fixed ROM partnumber list in listbox
set(handles.LoadSWC_partnumbers,'Enable','off')%Disable Load SWC.P/N button

% --- Executes on button press in Execute.
function Execute_Callback(hObject, eventdata, handles)
% hObject    handle to Execute (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global  SWC ROMTable FixedromTable ConstantListfile FixedROMPartnumber ROMPartnumber ConstantDBfile
assignin('base','FixedromTable',FixedromTable)%Stores FixedromTable in workspace
assignin('base','ROMTable',ROMTable)%Stores ROMTable in workspace

commanE1=strcat('ROMTable.',SWC);
command2=strcat('ROMTable.',ROMPartnumber); 
command3=strcat('FixedromTable.',FixedROMPartnumber);
command4=strcat('FixedromTable.',SWC);
ConstName=ROMTable.ConstantName;
FixConstName=FixedromTable.ConstantName;

Type=ROMTable.Type;
FixType=FixedromTable.Type;
swcy=evalin('base',commanE1);%Stores SWC column in array
Vehicle=evalin('base',command2);%Stores ROM partnumber column in array
FixVehicle=evalin('base',command3);% Stores Fixed partnumber column in array
Fixswcy=evalin('base',command4);% Stores Fixed SWC column in array
assignin('base','Fixswcy',Fixswcy)%Stores ROMTable in workspace
DataType=Type;
DataType= regexprep(DataType,'uchar8','uint8'); 
DataType= regexprep(DataType,'float32','single');
DataType= regexprep(DataType,'bool','boolean');
DataType= regexprep(DataType,'sint16','int16');
DataType= regexprep(DataType,'schar8','int8');
DataType= regexprep(DataType,'sint32','int32');
a=1;
for i=1:length(swcy)
    if  strcmp(swcy{i},'Y')&& strcmp(Vehicle{i},'Y')%Filters 'Y' in swc and ROM partnumber column
    Templatedata{a,1}=ConstName{i};
    Templatedata{a,2}=Type{i};
    Templatedata{a,3}=DataType{i};
    a=a+1;
    end    
end
Templatedata=array2table(Templatedata,'VariableNames',{'Parameter','type','typetrans'});
FixDataType=FixType;
FixDataType= regexprep(FixDataType,'uchar8','uint8'); 
FixDataType= regexprep(FixDataType,'float32','single');
FixDataType= regexprep(FixDataType,'bool','boolean');
FixDataType= regexprep(FixDataType,'sint16','int16');
FixDataType= regexprep(FixDataType,'schar8','int8');
FixDataType= regexprep(FixDataType,'sint32','int32');
a=1;
%Sorting required data
if ismember('Y',Fixswcy)
for i=1:length(Fixswcy)
    if  strcmp(Fixswcy{i},'Y')&& strcmp(FixVehicle{i},'Y')%Filters 'Y' in SWC and Fixed ROM partnumber column
    FixTemplatedata{a,1}=FixConstName{i};
    FixTemplatedata{a,2}=FixType{i};
    FixTemplatedata{a,3}=FixDataType{i};
    
    a=a+1;
    end    
end
FixTemplatedata=array2table(FixTemplatedata,'VariableNames',{'Parameter','type','typetrans'});
Templatedata = vertcat(Templatedata, FixTemplatedata);
end

unique(Templatedata,'stable');%Remove duplicated in filtered column
filename=strcat(SWC,'_ConstantDB_Check.xlsx');%Naming Report sheet

  

writetable(Templatedata,filename,'Sheet','Template','Range','H1','WriteVariableNames',true);%Writes filtered data in report
clc
[dummy, dummy, raw] = xlsread(ConstantListfile) ;%Reads ConstantListfile

table2=array2table(raw,'VariableNames',{'Parameter','block','block_type','intype','outType'});%Stores data in ConstantListfile in matlab table
parameter=table2.Parameter;%Extracts partnumber 
outType=table2.outType;
assignin('base','outType',outType)
a=1;

filename1=strcat('\',filename);
filename=strcat(pwd,filename1);
xlswrite(filename,raw,'Template')%Write data  in ConstantListfile into report

%writetable(Table2,filename,'Sheet','Template','Range','A1','WriteVariableNames',true);

exl = actxserver('excel.application');
exlWkbk = exl.Workbooks;
exlFile = exlWkbk.Open(filename);
exlFile.Activate
exlFile.Sheets.Item('Sheet1').Delete;
exlFile.Sheets.Item('Sheet2').Delete;
exlFile.Sheets.Item('Sheet3').Delete;
exlFile.Sheets.Item('Template').Activate
 for i=1:height(Templatedata)
         if ~ismember(Templatedata{i,1},parameter)  %Compares data between ConstantDB and Constantlistfile
     exlFile.Sheets.Item('Template').Range(sprintf('H%d',i+1)).Interior.ColorIndex = 3;
         end
         if ismember(Templatedata{i,1},parameter)
             for j=2:length(parameter)
                 if strcmp(Templatedata{i,1},parameter(j)) && strcmp(Templatedata{i,3},outType(j))==0
                    exlFile.Sheets.Item('Template').Range(sprintf('E%d',j)).Interior.ColorIndex = 3; 
                 end
             end
         end
 end
 for j=2:length(outType)
    if strcmp(outType(j),'boolean') || strcmp(outType(j),'int16')|| strcmp(outType(j),'int32')|| strcmp(outType(j),'int8')|| strcmp(outType(j),'uint16')|| strcmp(outType(j),'uint32')|| strcmp(outType(j),'uint8')|| strcmp(outType(j),'single')
 continue;
    else
        exlFile.Sheets.Item('Template').Range(sprintf('E%d',j)).Interior.ColorIndex = 3; 
    end
 end
exlFile.Save();
exlFile.Close();
exl.Quit;
exl.delete;

%Below code apply borders
 exl = actxserver('excel.application');
    exlWkbk = exl.Workbooks;
    exlFile = exlWkbk.Open(filename);
    exlFile.Activate

     exlFile.Sheets.Item('Template').Range('A1:E2000').Borders.Item('xlInsideHorizontal').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('A1:E2000').Borders.Item('xlInsideHorizontal').Weight = 2;
    exlFile.Sheets.Item('Template').Range('A1:E2000').Borders.Item('xlInsideVertical').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('A1:E2000').Borders.Item('xlInsideVertical').Weight = 2;
      exlFile.Sheets.Item('Template').Range('H1:J2000').Borders.Item('xlInsideHorizontal').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('H1:J2000').Borders.Item('xlInsideHorizontal').Weight = 2;
    exlFile.Sheets.Item('Template').Range('H1:J2000').Borders.Item('xlInsideVertical').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('H1:J2000').Borders.Item('xlInsideVertical').Weight = 2;
     exlFile.Sheets.Item('Template').Range('J1:J2000').Borders.Item('xlEdgeRight').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('J1:J2000').Borders.Item('xlEdgeRight').Weight = 2;
     exlFile.Sheets.Item('Template').Range('H1:H2000').Borders.Item('xlEdgeLeft').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('H1:H2000').Borders.Item('xlEdgeLeft').Weight = 2;
     exlFile.Sheets.Item('Template').Range('E1:E2000').Borders.Item('xlEdgeRight').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('E1:E2000').Borders.Item('xlEdgeRight').Weight = 2;
     exlFile.Sheets.Item('Template').Range('A1:E1').Interior.ColorIndex = 6;
     exlFile.Sheets.Item('Template').Range('H1:J1').Interior.ColorIndex = 6;
     
%      t=datetime('today');
%      day1=day(t,'name');
    
    exl.ActiveWindow.Zoom = 75;
    exlFile.Save();
    exlFile.Close();
    exl.Quit;
    exl.delete;
	%Below code writes information data in info sheet
    sheetname='info';
dat=datestr(now);
data1=cellstr(dat);
data2={'Date'};
xlswrite(char(filename),data1,sheetname,'B1');
xlswrite(char(filename),data2,sheetname,'A1');
t=datetime('today');
day1=day(t,'name');
data1=cellstr(day1);
data2={'Day'};
xlswrite(char(filename),data1,sheetname,'B2');
xlswrite(char(filename),data2,sheetname,'A2');
data1=cellstr(ROMPartnumber);
data2={'(ROM)'};
xlswrite(char(filename),data1,sheetname,'B3');
xlswrite(char(filename),data2,sheetname,'A3');
data1=cellstr(FixedROMPartnumber);
data2={'(FixedROM)'};
xlswrite(char(filename),data1,sheetname,'B4');
xlswrite(char(filename),data2,sheetname,'A4');
data1=cellstr(SWC);
data2={'SWC'};
xlswrite(char(filename),data1,sheetname,'B5');
xlswrite(char(filename),data2,sheetname,'A5');
data1=cellstr(ConstantListfile);
data2={'ConstList'};
xlswrite(char(filename),data1,sheetname,'B6');
xlswrite(char(filename),data2,sheetname,'A6');
data1=cellstr(ConstantDBfile);
data2={'dbfile'};
xlswrite(char(filename),data1,sheetname,'B7');
xlswrite(char(filename),data2,sheetname,'A7');
exl = actxserver('excel.application');
exlWkbk = exl.Workbooks;
exlFile = exlWkbk.Open(filename);
exlFile.Activate
exlFile.Sheets.Item('Template').Activate

exlFile.Save();
exlFile.Close();
exl.Quit;
exl.delete;
clc
 f = msgbox('Executed');
set(handles.Execute,'Enable','off')


% --- Executes on button press in End.
function End_Callback(hObject, eventdata, handles)
% hObject    handle to End (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
close(IO_ConstantCheckTool)


function csvtextedit_Callback(hObject, eventdata, handles)
% hObject    handle to csvtextedit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of csvtextedit as text
%        str2double(get(hObject,'String')) returns contents of csvtextedit as a double


% --- Executes during object creation, after setting all properties.
function csvtextedit_CreateFcn(hObject, eventdata, handles)
% hObject    handle to csvtextedit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function dbtextedit_Callback(hObject, eventdata, handles)
% hObject    handle to dbtextedit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of dbtextedit as text
%        str2double(get(hObject,'String')) returns contents of dbtextedit as a double


% --- Executes during object creation, after setting all properties.
function dbtextedit_CreateFcn(hObject, eventdata, handles)
% hObject    handle to dbtextedit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
