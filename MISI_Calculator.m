function varargout = MISI_Calculator(varargin)
% MISI_CALCULATOR MATLAB code for MISI_Calculator.fig
%      MISI_CALCULATOR, by itself, creates a new MISI_CALCULATOR or raises the existing
%      singleton*.
%
%      H = MISI_CALCULATOR returns the handle to a new MISI_CALCULATOR or the handle to
%      the existing singleton*.
%
%      MISI_CALCULATOR('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in MISI_CALCULATOR.M with the given input arguments.
%
%      MISI_CALCULATOR('Property','Value',...) creates a new MISI_CALCULATOR or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before MISI_Calculator_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to MISI_Calculator_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to plot_question MISI_Calculator

% Last Modified by GUIDE v2.5 10-Aug-2017 16:51:43

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @MISI_Calculator_OpeningFcn, ...
                   'gui_OutputFcn',  @MISI_Calculator_OutputFcn, ...
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


% --- Executes just before MISI_Calculator is made visible.
function MISI_Calculator_OpeningFcn(hObject, eventdata, handles, varargin)

%set default values for some calculator parameters
handles.output = hObject;
%for plotting flagged glucose curves in the calculator
handles.p=1;
%default for flagging criteria is one for all.
handles.flat=1;
handles.rebound=1;
handles.hypo=1;
%plot number to avoid over-writing files (initial value 1)
handles.std_plot_num=1;
handles.mod_plot_num=1;
%method of computing MISI emploted - will be used for plotting to specify
%cubic spline or linear interpolation
handles.plot_method=0;
handles.plot_std=0;
handles.plot_mod=0;
%generate empty arrays which will contain data 
handles.time_points=[0,30,60,90,120];
handles.glucose_data=[];
handles.insulin_data=[];
%logo 
logo_macsbio=imread('Macsbio_logo.png','BackgroundColor',[0.95,0.95,0.95]);
axes(handles.axes3)
image(logo_macsbio);
axis off;
% Update handles structure
guidata(hObject, handles);

% UIWAIT makes MISI_Calculator wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = MISI_Calculator_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in upload_glucose_file.
function upload_glucose_file_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%allows user to upload glucose OGTT data file.
%GUI will allow user to search computed directories for file. search will
%begin in directory from which the calculator is saved.
%Two copies of the uploaded data will be stored for use by other functions.
%glucose_original - will not be modified by any other function.
%glucose_data     - data used in calculation, may be modifed for unit
%                   conversion.
%--------------------------------------------------------------------------
%generates GUI for user to specify glucose file pathway
[file_name,file_path]=uigetfile('*.xlsx','Select glucose data file');
%constructs file pathway using output of uigetfile.
full_name=[file_path,file_name];
%reads in excel file specified by full_name and saves a version of it as
%glucose_original which will not be altered.
handles.glucose_original=xlsread(full_name);
%removes first row of input file - variable names.
handles.glucose_original(1,:)=[];
%saves a copy of the glucose_original as glucose_data
%glucose_data will be used for caclulations. (units may be converted)
handles.glucose_data=handles.glucose_original;
handles.pathway=file_path;
guidata(hObject,handles)

% --- Executes on button press in upload_insulin_file.
function upload_insulin_file_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%allows user to upload insulin OGTT data file.
%GUI will allow user to search computed directories for file. search will
%begin in directory from which the gluocse file was uploaded.
%Two copies of the uploaded data will be stored for use by other functions.
%insulin_original - will not be modified by any other function.
%insulin_data     - data used in calculation, may be modifed for unit
%                   conversion.
%--------------------------------------------------------------------------
%error message will pop up if glucose data has not be entered
if isempty(handles.glucose_data)
    uiwait(errordlg('please upload glucose file'))
    return
end
%generates GUI for user to specify insulin file pathway
[file_name,file_path]=uigetfile('*.xlsx','Select insulin data file',handles.pathway);
%constructs file pathway using output of uigetfile.
full_name=[file_path,file_name];
%reads in excel file specified by full_name and saves a version of it as
%insulin_original which will not be altered.
handles.insulin_original=xlsread(full_name);
%removes first row of input file - variable names.
handles.insulin_original(1,:)=[];
%saves a copy of the insulin_original as insulin_data
%insulin_data will be used for caclulations. (units may be converted)
handles.insulin_data=handles.insulin_original;
%gives error message if number of insulin measurements does not equal
%glucose measurements
if ~isequal(size(handles.insulin_data(1,2:end)),size(handles.glucose_data(1,2:end)));
    uiwait(errordlg('number of columns in glucose file does not match insulin file'))
    return
end
if ~isequal(size(handles.insulin_data(:,1)),size(handles.glucose_data(:,1)));
    uiwait(errordlg('number of rows in glucose file does not match insulin file'))
    return
end
guidata(hObject,handles)


function time_points_function_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%Allows user to enter sampling time point manually into calculator seperated 
%by a comma.
%[0,30,60,90,120]
%--------------------------------------------------------------------------
%will give error message if a glucose file has not been uploaded
if isempty(handles.glucose_data)
    uiwait(errordlg('please upload glucose file'))
    return
end
%will give error message if an insulin file has not been uploaded
if isempty(handles.insulin_data)
    uiwait(errordlg('please upload insulin file'))
    return
end
%gets user input from GUI in form of a string
tp_str=get(hObject,'String');
%converts this string to a vector of numbers
tp_num=str2num(char(tp_str));
%saves this vector of time points for use in all functions
handles.time_points=tp_num;
% will return error if the number of entered time points does not equal
%number of time points does not equal number of measurements per individual in glucose file. 
if ~isequal(size(tp_num),size(handles.glucose_data(1,2:end)));
    uiwait(errordlg('number of time points do not match glucose data'))
end
guidata(hObject,handles)

% --- Executes during object creation, after setting all properties.
function time_points_function_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in compute_std_misi.
function compute_std_misi_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%Will compute MISI using the original method proposed by Abdul-Ghani et
%al.(2007) and write output file to user specified directory. Will also save 
%figures for all flagged curves if save flagged plots button has been checked 
%(default is zero)
%--------------------------------------------------------------------------
%outputs error message if glucose data has not been uploaded
if isempty(handles.glucose_data)
    uiwait(errordlg('please upload glucose file'))
    return
end
%outputs error message if insulin data has not been uploaded
if isempty(handles.insulin_data)
    uiwait(errordlg('please upload insulin file'))
    return
end
%outputs error message if the number of time points does not equal the
%number of measurements per OGTT in glucose file.
if ~isequal(size(handles.time_points),size(handles.glucose_data(1,2:end)));
    uiwait(errordlg('number of time points does not match number of columns in glucose file'))
    return
end
%gives error message if number of insulin measurements does not equal
%glucose measurements
if ~isequal(size(handles.insulin_data(1,2:end)),size(handles.glucose_data(1,2:end)));
    uiwait(errordlg('number of columns in glucose file does not match insulin file'))
    return
end
if ~isequal(size(handles.insulin_data(:,1)),size(handles.glucose_data(:,1)));
    uiwait(errordlg('number of rows in glucose file does not match insulin file'))
    return
end
if numel(handles.glucose_data(2,2:end))<4
    uiwait(errordlg('too few time points in OGTT data to compute MISI'))
    return
end
if handles.time_points(end)<120
    uiwait(errordlg('final time point must be 120 minutes or later'))
    return
end
%creates progress bar for computation of MISI
h=waitbar(0,'Computing standard MISI please wait..');
%locks Compute standard MISI button so it cannot be selected again until
%computation is complete. (avoid perfoming multiple calculations
set(hObject,'Enable','Off');
handles.plot_method=1;
%creates arrays to store output, MISI values,dG/dt, I_bar, which glucose
%curves are flaggedm the reason for flagging and the suggeted values.
handles.MISI=zeros(numel(handles.glucose_data(:,1)),3);
handles.dG_dt=zeros(numel(handles.glucose_data(:,1)),3);
handles.I_bar=zeros(numel(handles.glucose_data(:,1)),2);
reason={};
problem=zeros(1);
suggest=zeros(numel(handles.glucose_data(:,1)),1)*NaN;
suggest_2=zeros(numel(handles.glucose_data(:,1)),1)*NaN;
hypo_list=zeros(numel(handles.glucose_data(:,1)),1);
%computes MISI for each individual (i)
for i=1:numel(handles.glucose_data(:,1))
    waitbar(i/numel(handles.glucose_data(:,1)));
    %stores glucose and insulin vector for individual i
    glu_data=handles.glucose_data(i,:);
    ins_data=handles.insulin_data(i,:);
    
    %checks for missing values in glucose or insulin vector.
    %if one or more values are missing individual is removed from analysis
    %this is recorded in the output file
    missing_g=sum(isnan(glu_data));
    missing_i=sum(isnan(ins_data));
    if missing_g>0 || missing_i>0
        handles.dG_dt(i,1)=glu_data(1);
        handles.dg_dt(i,2)=NaN;
        handles.dg_dt(i,3)=NaN;
        handles.I_bar(i,1)=glu_data(1);
        handles.I_bar(i,1)=NaN;
        handles.MISI(i,1)=glu_data(1);
        handles.MISI(i,2)=NaN;
        reason{i}='missing value(s)';
    else
    %if no missing values are found MISI is computed.
    
    %dG/dt is computed using line of best fit local this function fits the
    %line of best fit from peak to nadir of input data and also fileters
    %for user specified flagging criteria (flat==1,rebound==1, and hypo==1
    %by default)      
        
        dg_dt=line_of_best_fit_local(glu_data(2:end),handles.time_points,handles.flat,handles.rebound,handles.hypo);
   %computes I_bar as mean of given insulin data
        i_bar=mean(ins_data(2:end));
   %Computes MISI
        misi=abs(dg_dt(1))/i_bar;
         
   %Writes data identifier/ID number for individual i to all output arrays
        handles.dG_dt(i,1)=glu_data(1);
        handles.dg_dt(i,2)=dg_dt(1);
        handles.dg_dt(i,3)=dg_dt(2);
        handles.I_bar(i,1)=glu_data(1);
        handles.I_bar(i,1)=i_bar;
        handles.MISI(i,1)=glu_data(1);  
   

   %wites output of MISI to all data files, and filters for flagged glucose
   %curves
        if dg_dt(1)==10;
            %curves with a peak at 120 minutes - no suggested value computed
            handles.MISI(i,2)=NaN;
            reason{i}='peak at 120 minutes';
            problem(i)=i;
            %suggest(i)=NaN;
        elseif dg_dt(1)==20;
            %curves in which the glucose peak is less than 0.5 mmol/L
            %greater than the fasting value - suggested value computed
            handles.MISI(i,2)=NaN;
            reason{i}='flat glucose curve';
            problem(i)=i;
            a_dg_dt=line_of_best_fit_rebound(glu_data(2:end),handles.time_points);
            suggest(i)=1000*abs(a_dg_dt(1))/i_bar;
        elseif dg_dt(1)==30;
            %curves which are classed as having a large glucose rebound
            %suggested value computed using both nadir and global minimum
            handles.MISI(i,2)=NaN;
            reason{i}='large gluocse rebound';
            problem(i)=i;
            dg_dt_g=line_of_best_fit_rebound(glu_data(2:end),handles.time_points);
            suggest(i)=1000*abs(dg_dt_g(1))/i_bar;
            dg_dt_g_2=line_of_best_fit_max_min(glu_data(2:end),handles.time_points);
            suggest_2(i)=1000*abs(dg_dt_g_2(1))/i_bar;
        elseif dg_dt(1)==40;
            %curves in which the glucose concentration falls below 3.5 mmol/L
            %value suggested
            handles.MISI(i,2)=NaN;
            reason{i}='hypoglycaemia';
            problem(i)=i;
            a_dg_dt=line_of_best_fit_rebound(glu_data(2:end),handles.time_points);
            suggest(i)=1000*abs(a_dg_dt(1))/i_bar;
            hypo_list(i)=1;
        else
            %If curve has not been flagged MISI value is suggested
            handles.MISI(i,2)=1000*misi;
            reason{i}=NaN;
            %suggest(i)=NaN;
        end
    end
end
%close computation progress bar
close(h)
%open file saving progress bar
h=waitbar(0,'Saving files....');
%writes identifiers and suggested values for flagged gluocse curves to be used by other commands
handles.problem=problem(problem>0);
handles.reason=reason;
handles.suggest_2=suggest_2;
%allows user to specify directory for output file. by default it will open
%search in directory from which glucose and insulin files were uploaded
out_directory=uigetdir(handles.pathway,'Please select directory in which to write output file');
out_name=[out_directory,'\','standard_MISI_score','(',num2str(handles.std_plot_num),')','.xlsx'];
%writes output to excel file
headings={'ID number','MISI','Reason for exclusion','suggested value','Using global minimum'};
xlswrite(out_name,headings,1,'A1');
xlswrite(out_name,handles.MISI,1,'A2');
xlswrite(out_name,reason',1,'C2');
xlswrite(out_name,suggest,1,'D2');
xlswrite(out_name,suggest_2,1,'E2');
%if save plots was checked will generate and save plot for each flagged
%glucose curve
if handles.plot_std==1
    for p=1:numel(handles.problem)
        waitbar(p/numel(handles.problem));
        time=0:1:handles.time_points(end);
        rem=handles.problem(p);
        glu_rem=handles.glucose_data(rem,:);
        ins_rem=handles.insulin_data(rem,:);
        message=['individual ',num2str(handles.glucose_data(rem,1))];
        %uses line_of_best_fit_rebound rather than line_of_best_fit_local
        %function used above as this will provide a suggested value for
        %dG/dt
        dg_dt_rem=line_of_best_fit_rebound(handles.glucose_data(rem,2:end),handles.time_points);
        dg_dt_line=dg_dt_rem(1).*time+dg_dt_rem(2);
        p_fig=figure('visible','off');
        subplot(2,1,1)
        plot(handles.time_points,glu_rem(2:end));

        hold on;
        l1=plot(handles.time_points,glu_rem(2:end),'rx');


        title_mess=[message,handles.reason{rem},'(standard MISI)'];
        title(title_mess);

        l2=plot(time,dg_dt_line);
        axis([0,handles.time_points(end),0,max(glu_rem(2:end))+1]);
        legend([l1,l2],'glucose measurements','suggested dG/dt','Location','southeast','FontSise',6);
        %if the curve was falgged due to a large rebound the plot will also
        %display dG/dt computed from glucose peak to global minimum using
        %line_of_best_fit_max_min function
        if ~isnan(suggest_2(rem))
            dg_dt_max=line_of_best_fit_max_min(handles.glucose_data(rem,2:end),handles.time_points);
            dg_dt_max_line=dg_dt_max(1).*time+dg_dt_max(2);
            l3=plot(time,dg_dt_max_line);
            legend([l1,l2,l3],'glucose measurements','suggested dG/dt','dG/dt using global min','Location','southeast','FontSise',6);
        end
        %if the hypoglycaemia flagging option is checked the figures will
        %also contain a dashed line indicating 3.5 mmol/L
        if handles.hypo==1
            if hypo_list(rem)==1;
                l4=refline(0,3.5);
                l4.Color='k';
                l4.LineStyle=':';
                legend([l1,l2,l4],'glucose measurements','suggested dG/dt','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
                if ~isnan(suggest_2(rem))
                    legend([l1,l2,l3,l4],'glucose measurements','suggested dG/dt','dG/dt using global min','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
                end
            end
        end
        
        hold off
        xlabel('time (mins)')
        ylabel('plasma glucose (mmol/l)')

        subplot(2,1,2)
        plot(handles.time_points,ins_rem(2:end));
        hold on;    
        plot(handles.time_points,ins_rem(2:end),'rx');
        axis([0,handles.time_points(end),min(ins_rem(2:end))-10,max(ins_rem(2:end))+10]);
        hold off
        xlabel('time (mins)')
        ylabel('plasma insulin (pmol/l)')
        %saves figure to the specified out-put directory
        file_name=[out_directory,'\','std',message,'',handles.reason{rem},'(',num2str(handles.std_plot_num),')','.png'];
        saveas(p_fig,file_name);
    end
end
%close progress bar
close(h)
%stores output directory for use by other funtions (saving plot)
handles.out_directory=out_directory;
%resets the value of p
handles.p=1;
%will increase the file out put number for standard comutation by one to
%avoid overwriting of files for further computation of MISI
handles.std_plot_num=handles.std_plot_num+1;
%re-enables the button once calculation is complete
set(hObject,'Enable','On');
guidata(hObject,handles)

% --- Executes on button press in compute_misi_mod.
function compute_misi_mod_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%Will compute MISI using the modified cubic spline method proposed by O'Donovan et
%al.(2017) and write output file to user specified directory. Will also save 
%figures for all flagged curves if save flagged plots button has been checked
%(default is unchecked)
%--------------------------------------------------------------------------

%gives error message if no glucose file has been uploaded
if isempty(handles.glucose_data)
    uiwait(errordlg('please upload glucose file'))
    return
end
%gives error message if no insulin file has been uploaded
if isempty(handles.insulin_data)
    uiwait(errordlg('please upload insulin file'))
    return
end
%gives error message if number of entered time points does not equal
%glucose measurements
if ~isequal(size(handles.time_points),size(handles.glucose_data(1,2:end)));
    uiwait(errordlg('number of time points does not match number of columns in glucose file'))
    return
end
%gives error message if number of insulin measurements does not equal
%glucose measurements
if ~isequal(size(handles.insulin_data(1,2:end)),size(handles.glucose_data(1,2:end)));
    uiwait(errordlg('number of columns in glucose file does not match insulin file'))
    return
end
if ~isequal(size(handles.insulin_data(:,1)),size(handles.glucose_data(:,1)));
    uiwait(errordlg('number of rows in glucose file does not match insulin file'))
    return
end
if numel(handles.glucose_data(2,2:end))<4
    uiwait(errordlg('too few time points in OGTT data to compute MISI'))
    return
end
if handles.time_points(end)<120
    uiwait(errordlg('final time point must be 120 minutes or later'))
    return
end
%locks compute modified MISI button until calculation is complete
set(hObject,'Enable','Off');
%generates progress bar for calculation of modified MISI
h=waitbar(0,'Computing modified MISI please wait....');
%specifes method of cacluation for plotting of glucose curves (for use in
%plot flagged function)
handles.plot_method=2;
%generates arrays for storing results for use in other function
handles.MISI=zeros(numel(handles.glucose_data(:,1)),3);
handles.dG_dt=zeros(numel(handles.glucose_data(:,1)),3);
handles.I_bar=zeros(numel(handles.glucose_data(:,1)),2);
reason={};
suggest=zeros(numel(handles.glucose_data(:,1)),1)*NaN;
suggest_2=zeros(numel(handles.glucose_data(:,1)),1)*NaN;
problem=zeros(1);
hypo_list=zeros(numel(handles.glucose_data(:,1)),1);
%computes modified MISI for each individual i
for i=1:numel(handles.glucose_data(:,1))
    waitbar(i/numel(handles.glucose_data(:,1)));
    %saves glucose and insulin data for indivdual i
    glu_data=handles.glucose_data(i,:);
    ins_data=handles.insulin_data(i,:);
    %tests for any missing values. If missing values are found MISI is not
    %computed and this is noted in the output file
    missing_g=sum(isnan(glu_data));                                                                        
    missing_i=sum(isnan(ins_data));
    if missing_g>0 || missing_i>0
        handles.dG_dt(i,1)=glu_data(1);
        handles.dg_dt(i,2)=NaN;
        handles.dg_dt(i,3)=NaN;
        handles.I_bar(i,1)=glu_data(1);
        handles.I_bar(i,1)=NaN;
        handles.MISI(i,1)=glu_data(1);
        handles.MISI(i,2)=NaN;
        reason{i}='missing value(s)';
    else
    %if no missing values are found modified MISI is computed
    %firstly both the glucose and insulin data are splined. Fictional
    %fasting time points are added at -30,-15, and -7 mins to enforce a
    %steady state at 0 min.
        g_spline=spline([-30,-15,-7,handles.time_points],[glu_data(2),glu_data(2),glu_data(2),glu_data(2:end)],-30:1:handles.time_points(end));
        g_data=g_spline(31:end);
        i_spline=spline([-30,-15,-7,handles.time_points],[ins_data(2),ins_data(2),ins_data(2),ins_data(2:end)],-30:1:handles.time_points(end));
    %dG/dt and I_bar are computed on cubic spline inferred glucose and insulin curves. dG/dt is computed using
    %line_of_best_fit_local which fits line from peak to nadir and filters
    %for all user specifiec flagged glucose curves. (flat, large rebounds
    %and hypoglycaemia)
        dg_dt=line_of_best_fit_local(g_data,0:1:handles.time_points(end),handles.flat,handles.rebound,handles.hypo);
   
        i_bar=mean(i_spline(31:end));
   
        misi=abs(dg_dt(1))/i_bar;
          
   %identifier/ID number for individual i is written to all output files
        handles.dG_dt(i,1)=glu_data(1);
        handles.dg_dt(i,2)=dg_dt(1);
        handles.dg_dt(i,3)=dg_dt(2);
        handles.I_bar(i,1)=glu_data(1);
        handles.I_bar(i,1)=i_bar;
        handles.MISI(i,1)=glu_data(1);  
   

   %filters output for line_of_best_fit_local for flagged gluocse curves
   %and computes suggested modified MISI values where necessary
        if dg_dt(1)==10;
            %gluocse peak at 120 mins - modified MISI not computed - no suggested
            %value
            handles.MISI(i,2)=NaN;
            reason{i}='peak at 120 minutes';
            problem(i)=i;
        elseif dg_dt(1)==20;
            %flat glucose curve - modified MISI not computed - suggested value
            %computed from peak to nadir
            handles.MISI(i,2)=NaN;
            reason{i}='flat glucose curve';
            problem(i)=i;
            a_dg_dt=line_of_best_fit_rebound(g_data,0:1:handles.time_points(end));
            suggest(i)=1000*abs(a_dg_dt(1))/i_bar;
        elseif dg_dt(1)==30;
            %large glucose rebound - modified MISI not computed - suggested values
            %computed from peak to nadir and peak to global minimum
            handles.MISI(i,2)=NaN;
            reason{i}='large gluocse rebound';
            problem(i)=i;
            dg_dt_g=line_of_best_fit_rebound(g_data,0:1:handles.time_points(end));
            suggest(i)=1000*abs(dg_dt_g(1))/i_bar;
            dg_dt_max=line_of_best_fit_max_min(g_data,0:1:handles.time_points(end));
            suggest_2(i)=1000*abs(dg_dt_max(1))/i_bar;
        elseif dg_dt(1)==40;
            %hypoglycaemic  -  modified MISI not computed - suggested value computed
            %from peak to nadir
            handles.MISI(i,2)=NaN;
            reason{i}='hypoglycaemia';
            problem(i)=i;
            a_dg_dt=line_of_best_fit_rebound(g_data(2:end),0:1:handles.time_points(end));
            suggest(i)=1000*abs(a_dg_dt(1))/i_bar;
            hypo_list(i)=1;
        else
            %curve no flagged - modified MISI computed
            handles.MISI(i,2)=1000*misi;
            reason{i}=NaN;
        end    
    end
end
%close compute modified MISI progress bar
close(h)
%generates progress bar for file saving
h=waitbar(0,'Saving files');
%saves all information regarding flagged glucose curves for use in other
%functions (plot flagged glucose curves)
handles.problem=problem(problem>0);
handles.reason=reason;
handles.suggest_2=suggest_2;
%asks user to specify directory to which all output files will be written
out_directory=uigetdir(handles.pathway,'Please select directory in which to wirte output file');
out_name=[out_directory,'\','modified_MISI_score','(',num2str(handles.mod_plot_num),')','.xlsx'];
%wirtes results to modified MISI output excel file
headings={'ID number','MISI','Reason for exclusion','suggested value','Using global minimum'};
xlswrite(out_name,headings,1,'A1');
xlswrite(out_name,handles.MISI,1,'A2');
xlswrite(out_name,reason',1,'C2');
xlswrite(out_name,suggest,1,'D2');
xlswrite(out_name,suggest_2,1,'E2');
%if save plots box has been checked, generates and saves figures of glucose and insulin curves
%for all flagged individuals
if handles.plot_mod==1
    for p=1:numel(handles.problem)
        waitbar(p/numel(handles.problem));
        time=0:1:handles.time_points(end);
        rem=handles.problem(p);
        glu_rem=handles.glucose_data(rem,:);
        ins_rem=handles.insulin_data(rem,:);
        %generates spline of glucose and insulin data inforcing steady
        %state assumption
        g_spline=spline([-30,-15,-7,handles.time_points],[glu_rem(2),glu_rem(2),glu_rem(2),glu_rem(2:end)],-30:1:handles.time_points(end));
        g_data=g_spline(31:end);
        i_spline=spline([-30,-15,-7,handles.time_points],[ins_rem(2),ins_rem(2),ins_rem(2),ins_rem(2:end)],-30:1:handles.time_points(end));
        message=['individual ',num2str(handles.glucose_data(rem,1))];
        %computes dG/dt from peak to nadir using line_of_best_fit_rebound to display suggested
        %MISI value.
        dg_dt_rem=line_of_best_fit_rebound(g_data,time);
        dg_dt_line=dg_dt_rem(1).*time+dg_dt_rem(2);
        p_fig=figure('visible','off');
        subplot(2,1,1)
        plot(0:1:120,g_data);

        hold on;
        l1=plot(handles.time_points,glu_rem(2:end),'rx');

        title_mess=[message,handles.reason{rem},'(modified MISI)'];
        title(title_mess);

        l2=plot(time,dg_dt_line);
        axis([0,120,0,max(glu_rem(2:end))+1]);
        legend([l1,l2],'glucose measurements','suggested dG/dt','Location','southeast');
        %for curves flagged for large glucose rebound dG/dt is also
        %computed from peak to minimum
        if ~isnan(suggest_2(rem))
            dg_dt_max=line_of_best_fit_max_min(g_data,time);
            dg_dt_max_line=dg_dt_max(1).*time+dg_dt_max(2);
            l3=plot(time,dg_dt_max_line);
            legend([l1,l2,l3],'glucose measurements','suggested dG/dt','dG/dt using global min','Location','southeast','FontSise',6);
        end
        %if hypoglycaemia is a checked flagging option figures will also
        %inclued a dashed line at 3.5 mmol/L to indicate hypoglycaemia
        if handles.hypo==1
            if hypo_list(rem)==1;
                l4=refline(0,3.5);
                l4.Color='k';
                l4.LineStyle=':';
                legend([l1,l2,l4],'glucose measurements','suggested dG/dt','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
                if ~isnan(suggest_2(rem))
                    legend([l1,l2,l3,l4],'glucose measurements','suggested dG/dt','dG/dt using global min','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
                end     
            end
        end
        hold off
        xlabel('time (mins)')
        ylabel('plasma glucose (mmol/l)')

        subplot(2,1,2)
        plot(time,i_spline(31:end));
        hold on;    
        plot(handles.time_points,ins_rem(2:end),'rx');
        axis([0,time(end),min(i_spline(31:end))-10,max(i_spline(31:end))]);
        hold off
        xlabel('time (mins)')
        ylabel('plasma insulin (pmol/l)')
        %saves figures to output directory specified by user.
        file_name=[out_directory,'\','mod ',message,'',handles.reason{rem},'(',num2str(handles.mod_plot_num),')','.png'];
        saveas(p_fig,file_name);
    end
end
%Saves user specified output directory for use by other functions
handles.out_directory=out_directory;
close(h);
%resets value of p for use by other functions
handles.p=1;
%increases file plot number by one to avoid over-writing files.
handles.mod_plot_num=handles.mod_plot_num+1;    
set(hObject,'Enable','On');
guidata(hObject,handles)


% --- Executes on button press in plot_flagged_curves.
function plot_flagged_curves_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%will plot flagged glucose and insulin curves for individual p as identified 
%by most recent calculation of MISI (either standard or modified methods).
%each time plot button is selected on GUI p will increase by one. 
%--------------------------------------------------------------------------

%will return error message if MISI has no been computed.
if handles.plot_method==0
    uiwait(errordlg('please compute MISI'))
    return
end
%defines necessay data to compute MISI for pth individual in flagged list    
time=0:1:handles.time_points(end);
rem=handles.problem(handles.p);
glu_rem=handles.glucose_data(rem,:);
ins_rem=handles.insulin_data(rem,:);
message=['individual ',num2str(glu_rem(1))];

%if MISI was computed using the standard method plots will use linear
%interpolation between points and suggested MISI values will also be
%computed using the standard method
if handles.plot_method==1
    %plots data points and dG/dt computed using standard method computed
    %from peak to nadir.
    test_hypo=line_of_best_fit_local(glu_rem(2:end),handles.time_points,1,1,1);
    dg_dt_rem=line_of_best_fit_rebound(glu_rem(2:end),handles.time_points);
    dg_dt_line=dg_dt_rem(1).*time+dg_dt_rem(2);
    axes(handles.axes1);
    plot(handles.time_points,glu_rem(2:end));

    hold on;
    l1=plot(handles.time_points,glu_rem(2:end),'rx');


    title({message,handles.reason{rem}});

    l2=plot(time,dg_dt_line);
    axis([0,time(end),0,max(glu_rem(2:end))+1]);
    legend([l1,l2],'glucose measurements','suggested dG/dt','Location','southeast');
    %if glucose curve was flagged due to large glucose rebound suggested
    %MISI value will also be computed from peak to global minimum. 
    if ~isnan(handles.suggest_2(rem))
        dg_dt_max=line_of_best_fit_max_min(glu_rem(2:end),handles.time_points);
        dg_dt_max_line=dg_dt_max(1).*time+dg_dt_max(2);
        l3=plot(time,dg_dt_max_line);
        legend([l1,l2,l3],'glucose measurements','suggested dG/dt','dG/dt using global min','Location','southeast','FontSise',6);
    end
    %if glucose curves were also flagged for hypoglycaemia a dashed line
    %will appear on glucose plots to indicate 3.5 mmol/L
    if handles.hypo==1
        if test_hypo(1)==40;
            l4=refline(0,3.5);
            l4.Color='k';
            l4.LineStyle=':';
            legend([l1,l2,l4],'glucose measurements','suggested dG/dt','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            if ~isnan(handles.suggest_2(rem))
                legend([l1,l2,l3,l4],'glucose measurements','suggested dG/dt','dG/dt using global min','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            end
        end
    end
    hold off
    xlabel('time (mins)')
    ylabel('plasma glucose (mmol/l)')

    axes(handles.axes2);
    plot(handles.time_points,ins_rem(2:end));
    hold on;
    plot(handles.time_points,ins_rem(2:end),'rx');
    axis([0,time(end),min(ins_rem(2:end))-10,max(ins_rem(2:end))+10]);
    hold off
    xlabel('time (mins)')
    ylabel('plasma insulin (pmol/l)')
%if modifed method was used to compute MISI plots will display cubic spline
%of both glucose and insulin data and dG/dt will be computed using modified
%method
elseif handles.plot_method==2
    %cubic splines of both insulin and glucose data are generated enforcing
    %steady state before 0 mins
    g_spline=spline([-30,-15,-7,handles.time_points],[handles.glucose_data(rem,2),handles.glucose_data(rem,2),handles.glucose_data(rem,2),handles.glucose_data(rem,2:end)],-30:1:handles.time_points(end));
    i_spline=spline([-30,-15,-7,handles.time_points],[handles.insulin_data(rem,2),handles.insulin_data(rem,2),handles.insulin_data(rem,2),handles.insulin_data(rem,2:end)],-30:1:handles.time_points(end));
    %dG/dt is computed using line_of_best_fit_rebound function which
    %fits line from peak to nadir
    dg_dt_rem=line_of_best_fit_rebound(g_spline(31:end),time);
    dg_dt_line=dg_dt_rem(1).*time+dg_dt_rem(2);
    test_hypo=line_of_best_fit_local(g_spline(31:end),time,1,1,1);
    axes(handles.axes1);
    plot(time,g_spline(31:end));
    hold on;
    l1=plot(handles.time_points,handles.glucose_data(rem,2:end),'rx');

    title({message,handles.reason{rem}});

    l2=plot(time,dg_dt_line);
    axis([0,time(end),0,max(glu_rem(2:end))+1]);
    legend([l1,l2],'glucose measurements','suggested dG/dt','Location','southeast');
    %if glucose curve was flagged due to a large glucose rebound dG/dt will
    %also be computed from glucose peak to global minimum and displayed
    if ~isnan(handles.suggest_2(rem))
        dg_dt_max=line_of_best_fit_max_min(g_spline(31:end),time);
        dg_dt_max_line=dg_dt_max(1).*time+dg_dt_max(2);
        l3=plot(time,dg_dt_max_line);
        legend([l1,l2,l3],'glucose measurements','suggested dG/dt','dG/dt using global min','Location','southeast','FontSise',6);
    end
    %if hypoglycaemia was a flagging option glucose will contain a dashed
    %line which will indicate 3.5 mmol/L
    if handles.hypo==1
        if test_hypo(1)==40;
            l4=refline(0,3.5);
            l4.Color='k';
            l4.LineStyle=':';
            legend([l1,l2,l4],'glucose measurements','suggested dG/dt','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            if ~isnan(handles.suggest_2(rem))
                legend([l1,l2,l3,l4],'glucose measurements','suggested dG/dt','dG/dt using global min','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            end
        end
    end
    hold off
    xlabel('time (mins)')
    ylabel('plasma glucose (mmol/l)')


    axes(handles.axes2);
    plot(time,i_spline(31:end));
    hold on;
    plot(handles.time_points,ins_rem(2:end),'rx');
    axis([0,time(end),min(i_spline(31:end))-10,max(i_spline(31:end))+10]);
    hold off
    xlabel('time (mins)')
    ylabel('plasma insulin (pmol/l)')
end

%once plot is displayed p is increased by one to plot the next individual
handles.p=handles.p+1;
%once plot for final flagged individual has been displayed (p>number of
%flagged individuals) an error message will be displayed to inform user curves for all
%individuals have been displayed
if handles.p>numel(handles.problem)
    uiwait(errordlg('All flagged curves have been plotted'))
end

guidata(hObject,handles)

% --- Executes on button press in save_std.
function save_std_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%check box to save figures of glucose and insulin curves for all flagged
%individuals when MISI is computed using the standard method.
%--------------------------------------------------------------------------
%if box is checked plot_std==1, if box is unchecked plot_std==0 (default)
%--------------------------------------------------------------------------
plot_std=get(hObject,'Value');
handles.plot_std=plot_std;
%if box is check a warning message will be displayed to warn user that
%saving figures for each flagged glucose curve will take some time.
if plot_std==1
    msgbox('Caution; May take longer to compute when saving figures')
end
guidata(hObject,handles)

% --- Executes on button press in save_mod.
function save_mod_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%check box to save figures of glucose and insulin curves for all flagged
%individuals when MISI is computed using the modified method.
%--------------------------------------------------------------------------
%if box is checked plot_mod==1, if box is unchecked plot_mod==0 (default)
%--------------------------------------------------------------------------
plot_mod=get(hObject,'Value');
handles.plot_mod=plot_mod;
%if box is check a warning message will be displayed to warn user that
%saving figures for each flagged glucose curve will take some time.
if plot_mod==1
    msgbox('Caution; May take longer to compute when saving figures')
end
guidata(hObject,handles);

% --- Executes on button press in question_upload.
function question_upload_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%if help button/question button in step 1 (uploading data files) is
%selceted the following message will be displayed
%--------------------------------------------------------------------------
options.Interpreter='tex';
options.WindowStyle='modal';
msgbox({'\bf Uploading data files \rm ','','Glucose and insulin OGTT data should be stored in separate files in .xlsx format. The first row of each file may contain variable names (this row will not be used in the computation of MISI). The remaining rows should contain the OGTT data for each individual, with the first column containing an identifier (ID number) and the remaining columns containing the glucose/insulin measurements for each sampled time point of the OGTT. Files may be uploaded with missing values as the calculator will filter out OGTT data with one or more missing values (this will be noted in the output file). Sample glucose and insulin data files can be found at website','','\bf Time points \rm','','Sampling time points should be entered directly into the window of the calculator separated by a comma as follows;',' 0,30,60,90,120 ','followed by enter. If the number of entered time points do not match the number of columns in the uploaded glucose and/or insulin files the calculator will not compute MISI and you will receive an error message.'},'Help',options);

% --- Executes on button press in question_flag_option.
function question_flag_option_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%if help button/question button in step 4 (flagging criteria) is selected
%the following message will be displayed
%--------------------------------------------------------------------------
options.Interpreter='tex';
options.WindowStyle='modal';
msgbox({'\bf Flagging options \rm', '',' Allows user to specify criteria for which glucose curves will be flagged for manual evaluation of computed/suggested MISI values.','- peak at 120 minutes (not optional)','- flat glucose curves: peak is less than 0.5 mmol/L greater than the fasting value','- large glucose rebound: rebound in glucose concentration following nadir is greater than 0.5mmol/L. Particularly an issue in frequently sampled OGTT data. Calculator will suggest a MISI value computed from peak to nadir and a second value computed from peak to global minimum','- hypoglycaemia: curves where glucose concentration falls below 3.5 mmol/L','Please refer to paper reference for more information.(see \it About \rm )'},'Help',options);

% --- Executes on button press in question_std_misi.
function question_std_misi_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%if help button/question button next to step 5 (compute standard MISI) is
%selected the following message will pop up
%--------------------------------------------------------------------------
options.Interpreter='tex';
options.WindowStyle='modal';
msgbox({'\bf Computing MISI - standard method \rm','','Will compute MISI using the original method as outlined by Abdul Ghani \it et al.\rm (2007). MISI will be computed in (\mumol.min)/pmol with mean insulin calculated using all supplied time points.','','\it Output file \rm','',' The user will be asked to specify the directory to which the output file(s) will be written. The first column of the MISI output file will contain the user supplied identifiers (ID numbers). The second column will contain the computed MISI value, this will be blank if the glucose curve was flagged or if there were missing values in the given OGTT data for that individual. The third column will contain the reason for a missing MISI value. The fourth column will contain a suggested MISI value computed, where possible, using the original method (no value will be suggested for missing OGTT values or peaks at 120 mins). The fifth column will contain a suggested MISI value computed using the global minimum rather than the nadir to fit dG/dt.','','\it Saving flagged glucose curves \rm', '','It is possible to save the plotted glucose and insulin curves along with dG/dt for the suggested MISI values for all flagged curves when computing MISI. This can be done by checking the save file box below the \it Compute standard MISI \rm or \it Compute modified MISI \rm buttons. All plots will be saved in the same directory specified by the user for the output file. \bf Caution! \rm  Saving all plots may increase the time take to compute MISI. It is also possible to save plots on visual inspection with the calculator.'},'Help',options);

% --- Executes on button press in question_mod_misi.
function question_mod_misi_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%if the help button/question button in step 5 (computed modified MISI) is
%selected the following message will pop up
%--------------------------------------------------------------------------
options.Interpreter='tex';
options.WindowStyle='modal';

msgbox({'\bf Computing MISI - modified method \rm', '','Will compute MISI using the modified cubic spline method as outlined by ODonovan \it et al. \rm (paper in preperation). dG/dt will be computed on a cubic spline of the supplied glucose data in \mumol/L/min with the cubic spline giving an improved prediction of the glucose peak and nadir. Mean insulin computed on a cubic spline of the supplied insulin data in pmol/L as this will account for unequally spaced sampling frequencies. The modified method has yielded an improved correlation with the steady state glucose infusion rate of the hyperinsulinemic euglycaemic clamp and is recommended for use when computing MISI on five or less time points or when unequally spaced sampling frequency was used during the OGTT.', '','\it Output file \rm','',' The user will be asked to specify the directory to which the output file(s) will be written. The first column of the MISI output file will contain the user supplied identifiers (ID numbers). The second column will contain the computed MISI value, this will be blank if the glucose curve was flagged or if there were missing values in the given OGTT data for that individual. The third column will contain the reason for a missing MISI value. The fourth column will contain a suggested MISI value computed, where possible, using the original method (no value will be suggested for missing OGTT values or peaks at 120 mins). The fifth column will contain a suggested MISI value computed using the global minimum rather than the nadir to fit dG/dt.','','\it Saving flagged glucose curves \rm', '','It is possible to save the plotted glucose and insulin curves along with dG/dt for the suggested MISI values for all flagged curves when computing MISI. This can be done by checking the save file box below the \it Compute standard MISI \rm or \it Compute modified MISI \rm buttons. All plots will be saved in the same directory specified by the user for the output file. \bf Caution! \rm  Saving all plots may increase the time take to compute MISI. It is also possible to save plots on visual inspection with the calculator.'},'Help',options);

% --- Executes on button press in flat.
function flat_Callback(hObject, eventdata, handles)
%-------------------------------------------------------------------------
%if box indicating flag is check or unchecked in flagging criteria check 
%panel this function will save this to GUI handles which can be found by
%all other functions.
%-------------------------------------------------------------------------
f=get(hObject,'Value');
if f==1
    handles.flat=1;
elseif f==0
    handles.flat=0;
end
guidata(hObject,handles);

% --- Executes on button press in rebound.
function rebound_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%if box indicating flaggind due to a large glucose rebound is checked or 
%unchecked in flagging criteria panel this function will save this to
%GUI handels for use by other functions in calculator.
%--------------------------------------------------------------------------
r=get(hObject,'Value');
if r==1
    handles.rebound=1;
elseif r==0
    handles.rebound=0;
end
guidata(hObject,handles)

% --- Executes on button press in hypo.
function hypo_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%if box indicating flagging due to hypoglycaemia in flagging criteria panel
%is checked or unchecked this function will save this change to
%GUI handels for use by other functions in calculator.
%--------------------------------------------------------------------------
h=get(hObject,'Value');
if h==1
    handles.hypo=1;
elseif h==0
    handles.hypo=0;
end
guidata(hObject,handles)

function slope=line_of_best_fit_rebound(data_g,data_t)
%-------------------------------------------------------------------------
%computes line of best fit from peak to nadir for a given vector of glucose
%values sampled at given time points.
%-------------------------------------------------------------------------
%data_g  - vector of glucose measurements [5.5,6.5,7.5,6.5,5.5]
%data_t  - vector of sampling time points [0,30,60,90,120]
%
%output  - [slope,intercept]
%-------------------------------------------------------------------------
%computes dG/dt from peak to nadir - does not filter for any flagging
%options. Will not find unique solution if peak occurs at final time point.
%--------------------------------------------------------------------------
%finds maximum point in glucose vector
[max_val,max_loc]=max(data_g);
%finds global minimum in gluocse vector after maximum
[min_val,min_loc]=min(data_g(max_loc:end));
%define new time and points vectors as all points between maximum and
%minimum inclusive.
time=data_t(max_loc:max_loc+min_loc-1); 
points=data_g(max_loc:max_loc+min_loc-1);


%determine if there is an earlier minimum before the globally detected one.
%this is acheived by testing for a peak occuring between the identified
%global maximumand minimum
if numel(data_g(max_loc:min_loc))>=3;
    [p_val,p_loc]=findpeaks(data_g(max_loc:max_loc+min_loc-1));
    if p_loc ~= 1;
        [min_val_2,min_loc_2]=min(data_g(max_loc:max_loc+p_loc-1));
        points=data_g(max_loc:max_loc+min_loc_2-1);
        time=data_t(max_loc:max_loc+min_loc_2-1);
    end
end




slope=polyfit(time,points,1);

function slope=line_of_best_fit_local(data_g,data_t,flat,rebound,hypo)
%-------------------------------------------------------------------------
%computes line of best fit from peak to nadir for a given vector of glucose
%values sampled at given time points. Filters for peaks at final time point
%and given flagging options
%-------------------------------------------------------------------------
%data_g  - vector of glucose measurements [5.5,6.5,7.5,6.5,5.5]
%data_t  - vector of sampling time points [0,30,60,90,120]
%flat    - logical: 1 - filter for flat curves, 0 - do not filter flat curves
%rebound - logical: 1 - filter for curves with rebound, 0 - do not filter
%hypo    - logical: 1 - filter for curves with a value less than 3.5 mmol/L
%                   0 - do not filter for curves with a value less than
%                       3.5mmol/l
%
%output  - [slope,intercept]
%        - [10,0] indicates a peak at final time point
%        - [20,0] indicates a flat glucose curve, peak less than 0.5 mmol/L
%                 greater than the fasting value.
%        - [30,0] indicates a glucose curve with a rebound larger than 0.5
%                 mmol/L greater than the nadir.
%        - [40,0] indicates hypoglycaemia, curves with a value less than
%                 3.5mmol/L
%-------------------------------------------------------------------------
%computes dG/dt from peak to nadir - filters for user specified flagging
%options. Will not find unique solution if peak occurs at final time point.
%--------------------------------------------------------------------------
%finds maximum point in glucose vector
[max_val,max_loc]=max(data_g);
%finds minimum in gluocse vector after maximum
[min_val,min_loc]=min(data_g(max_loc:end));
%define new time and points vectors as all points between maximum and
%minimum inclusive.
time=data_t(max_loc:max_loc+min_loc-1); 
points=data_g(max_loc:max_loc+min_loc-1);



%determine if there is an earlier minimum before the globally detected one.
%this is done by checking for the presence of a peak between the global
%maximum and minimum. 
if numel(data_g(max_loc:min_loc))>=3;
    [p_val,p_loc]=findpeaks(data_g(max_loc:max_loc+min_loc-1));
    if numel(p_loc)>0;
        [min_val_2,min_loc_2]=min(data_g(max_loc:max_loc+p_loc-1));
        points=data_g(max_loc:max_loc+min_loc_2-1);
        time=data_t(max_loc:max_loc+min_loc_2-1);
    end
end

%filters for flagging options
if max_loc==numel(data_g);
    %peak at 120 assigns a value of 10
    slope=[10,0];
elseif max_val-data_g(1)<=0.5;
    %curves with no signicant peak "flat curves"
    if flat==1
        slope=[20,0];
    else
        slope=polyfit(time,points,1);
    end
elseif sum(data_g<3.5)>=1;
    %hypoglycaemia
    if hypo==1;
        slope=[40,0];
    else
        slope=polyfit(time,points,1);
    end
elseif max_val==min_val
    %actual flat curves; this does not flag the curves but simply force them
    %to have the expected slope of 0
    slope=[0,max_val];
elseif numel(data_g(max_loc:min_loc))>=3;
    %finds curves which have a rebound and if the rebound is larger than
    %0.5 mmol/L flags them
    [p_val,p_loc]=findpeaks(data_g(max_loc:max_loc+min_loc-1));
    if numel(p_loc)>0
        if min_val_2-min_val<=0.5 && data_g(max_loc+min_loc_2)-data_g(max_loc+min_loc_2-1)<=0.5;
            slope=polyfit(time,points,1);
        else
            %rebound
            if rebound==1;
                slope=[30,0];
            else 
                slope=polyfit(time,points,1);
            end
        end
    else
        slope=polyfit(time,points,1);
    end
else 
    %if curve is not flagged for any reason computes line of best fit from
    %peak to nadir using polyfit function 
    slope=polyfit(time,points,1);
end

function slope=line_of_best_fit_max_min(data_g,data_t)
%-------------------------------------------------------------------------
%computes line of best fit from peak to global minimum for a given vector 
%of glucose values sampled at given time points.
%-------------------------------------------------------------------------
%data_g  - vector of glucose measurements [5.5,6.5,7.5,6.5,5.5]
%data_t  - vector of sampling time points [0,30,60,90,120]
%
%output  - [slope,intercept]
%--------------------------------------------------------------------------
%computes dG/dt from peak to global minimum - does not filter for user 
%specified flagging options. Will not find unique solution if peak occurs 
%at final time point.
%--------------------------------------------------------------------------
%finds maximum point in glucose vector
[max_val,max_loc]=max(data_g);
%finds minimum in gluocse vector after maximum
[min_val,min_loc]=min(data_g(max_loc:end));
%define new time and points vectors as all points between maximum and
%global minimum inclusive.
time=data_t(max_loc:max_loc+min_loc-1); 
points=data_g(max_loc:max_loc+min_loc-1);

slope=polyfit(time,points,1);


% --- Executes when selected object is changed in glucose_unit.
function glucose_unit_SelectionChangedFcn(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%converts glucose units to SI units mmol/L based on selection of button in
%glucose unit panel. Uses glucose_original to prevent overwriting of input
%data if incorrect units are selected
%--------------------------------------------------------------------------

%gets user specified measurement unit from GUI in form of a string
unit=get(hObject,'Tag');
if strcmp(unit,'mgdl');
    %if measurement units are mg/dL input values are divided by 18 to
    %convert to mmol/L
    handles.glucose_data=[handles.glucose_original(:,1),handles.glucose_original(:,2:end)/18];
elseif strcmp(unit,'mmoll')
    %if measurement units are mmol/L units remain unchanged
    %each time this button is selected the data will be copied from the
    %originally uploaded data to avoid over-writing of data with incorrect
    %unit conversions.
    handles.glucose_data=handles.glucose_original;
end
    
guidata(hObject,handles)


% --- Executes when selected object is changed in insulin_units.
function insulin_units_SelectionChangedFcn(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%converts glucose units to SI units pmol/L based on selection of button in
%insulin unit panel. Uses insulin_original to prevent over-writing of input
%data if incorrect units are selected
%--------------------------------------------------------------------------

%gets user specified insulin units from GUI in the form of a string
unit=get(hObject,'Tag');
if strcmp(unit,'uUml')
    %if measurment units are in uU/ml than the uploaded insulin
    %measurements are multiplied by 6.945 to convert them to pmol/L
    handles.insulin_data=[handles.insulin_original(:,1),handles.insulin_original(:,2:end).*6.945];
elseif strcmp(unit,'pmoll')
    %if measurement units are already in pmol/L the values remain unchanged
    handles.insulin_data=handles.insulin_original;
end
guidata(hObject,handles)


% --- Executes on button press in save_plot_onfly.
function save_plot_onfly_Callback(hObject, eventdata, handles)
%-------------------------------------------------------------------------
%This function allows user to save individual glucose and insulin plots 
%displayed on the calculator GUI at the time. The figure will be saved in 
%the output directory specified by the user during the calculation of MISI
%-------------------------------------------------------------------------

%specified individual to be saved (1-current value of P)
rem=handles.problem(handles.p-1);
%saves glucose and insulin data for this individual.
glu_rem=handles.glucose_data(rem,:);
ins_rem=handles.insulin_data(rem,:);
time=0:1:handles.time_points(end);
message=['individual ',num2str(glu_rem(1))];

%if MISI was computed using standard method
if handles.plot_method==1
    %compute suggested dG/dt value computed using stadard method from peak
    %to nadir
    dg_dt_rem=line_of_best_fit_rebound(glu_rem(2:end),handles.time_points);
    dg_dt_line=dg_dt_rem(1).*time+dg_dt_rem(2);
    p_fig=figure('visible','off');
    hypo_test=line_of_best_fit_local(glu_rem(2:end),handles.time_points,0,0,1);
    subplot(2,1,1)
    plot(handles.time_points,glu_rem(2:end));

    hold on;
    l1=plot(handles.time_points,glu_rem(2:end),'rx');


    title({message,handles.reason{rem}});

    l2=plot(time,dg_dt_line);
    axis([0,time(end),0,max(glu_rem(2:end))+1]);
    legend([l1,l2],'glucose measurements','suggested dG/dt','Location','southeast');
    %if curve was flagged for a large glucose rebound also plot dG/dt
    %computed from glucose peak to global minimum.
    if ~isnan(handles.suggest_2(rem))
        dg_dt_max=line_of_best_fit_max_min(glu_rem(2:end),handles.time_points);
        dg_dt_max_line=dg_dt_max(1).*time+dg_dt_max(2);
        l3=plot(time,dg_dt_max_line);
        legend([l1,l2,l3],'glucose measurements','suggested dG/dt','dG/dt using global min','Location','southeast','FontSise',6);
    end
    %if hypoglycaema is selected as a flagging option flot will include a
    %dashed line indicating 3.5 mmol/L
    if handles.hypo==1
        if hypo_test(1)==40;
            l4=refline(0,3.5);
            l4.Color='k';
            l4.LineStyle=':';
            legend([l1,l2,l4],'glucose measurements','suggested dG/dt','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            if ~isnan(handles.suggest_2(rem))
                legend([l1,l2,l3,l4],'glucose measurements','suggested dG/dt','dG/dt using global min','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            end
        end
    end
    hold off
    xlabel('time (mins)')
    ylabel('plasma glucose (mmol/l)')


    subplot(2,1,2)
    %also plots insulin
    plot(handles.time_points,ins_rem(2:end));
    hold on;
    plot(handles.time_points,ins_rem(2:end),'rx');
    axis([0,time(end),min(ins_rem(2:end))-10,max(ins_rem(2:end))+10]);
    hold off
    xlabel('time (mins)')
    ylabel('plasma insulin (pmol/l)')
    %saves figure to output directory specified by user when calculating
    %MISI
    file_name=[handles.out_directory,'\','mod ',message,'',handles.reason{rem},' manual(',num2str(handles.mod_plot_num),')','.png'];
    saveas(p_fig,file_name);
%if MISI was computed using the modified cubic spline method
elseif handles.plot_method==2
    %computes cubic spline of both glucose and insulin data enforcing
    %steady state critia before 0 min.
    g_spline=spline([-30,-15,-7,handles.time_points],[handles.glucose_data(rem,2),handles.glucose_data(rem,2),handles.glucose_data(rem,2),handles.glucose_data(rem,2:end)],-30:1:handles.time_points(end));
    i_spline=spline([-30,-15,-7,handles.time_points],[handles.insulin_data(rem,2),handles.insulin_data(rem,2),handles.insulin_data(rem,2),handles.insulin_data(rem,2:end)],-30:1:handles.time_points(end));
    %computes dG/dt from glucose peak to nadir using cubic spline of glucose data.
    dg_dt_rem=line_of_best_fit_rebound(g_spline(31:end),time);
    dg_dt_line=dg_dt_rem(1).*time+dg_dt_rem(2);
    
    hypo_test=line_of_best_fit_local(g_spline(31:end),time,0,0,1);
    p_fig=figure('visible','off');
    subplot(2,1,1)
    plot(time,g_spline(31:end));
    hold on;
    l1=plot(handles.time_points,handles.glucose_data(rem,2:end),'rx');

    title({message,handles.reason{rem}});

    l2=plot(time,dg_dt_line);
    axis([0,time(end),0,max(glu_rem(2:end))+1]);
    legend([l1,l2],'glucose measurements','suggested dG/dt','Location','southeast');
    %if glucose curve was flagged due to large glucose rebound plot will
    %also contain dG/dt computed from glucose peak to global minimum on
    %glucose spline.
    if ~isnan(handles.suggest_2(rem))
        dg_dt_max=line_of_best_fit_max_min(g_spline(31:end),time);
        dg_dt_max_line=dg_dt_max(1).*time+dg_dt_max(2);
        l3=plot(time,dg_dt_max_line);
        legend([l1,l2,l3],'glucose measurements','suggested dG/dt','dG/dt using global min','Location','southeast','FontSise',6);
    end
    %if hypoglycaemia was selected as a flagging criteria plot will contain
    %a dashed line indicating 3.5 mmol/L.
    if handles.hypo==1
        if hypo_test(1)==40;
            l4=refline(0,3.5);
            l4.Color='k';
            l4.LineStyle=':';
            legend([l1,l2,l4],'glucose measurements','suggested dG/dt','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            if ~isnan(handles.suggest_2(rem))
                legend([l1,l2,l3,l4],'glucose measurements','suggested dG/dt','dG/dt using global min','hypoglycaemia 3.5 mmol/l','Location','southeast','FontSise',6);
            end
        end
    end
    hold off
    xlabel('time (mins)')
    ylabel('plasma glucose (mmol/l)')


    subplot(2,1,2)
    %cubic spline of insulin data is also plotted.
    plot(time,i_spline(31:end));
    hold on;
    plot(handles.time_points,ins_rem(2:end),'rx');
    axis([0,time(end),min(i_spline(31:end))-10,max(i_spline(31:end))+10]);
    hold off
    xlabel('time (mins)')
    ylabel('plasma insulin (pmol/l)')
    %figure is saved to output direcotry specified by user in calculation
    %of modified MISI.
    file_name=[handles.out_directory,'\','mod ',message,'',handles.reason{rem},' manual(',num2str(handles.mod_plot_num),')','.png'];
    saveas(p_fig,file_name);
end


% --- Executes on button press in about.
function about_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%displays message on selection of about button
%--------------------------------------------------------------------------
options.Interpreter = 'tex';
options.WindowStyle='modal';


msgbox({'\bf MISI Calculator 1.0 \rm','','Calculator allows for the computation of the muscle insulin sensitivity index on user supplied oral glucose tolerance test data.','','\it Features\rm','--------------------------------------------------------------------------','- Upload glucose and insulin OGTT data in ''.xlsx'' format','- Inbuilt filtering out of missing data','- Conversion of units for standardised calculation of MISI','- User specified flagging of problematic glucose curves for manual evaluation','- Computation of MISI using the original method or modified cubic spline method.','- Plotting of flagged glucose curves for manual inspection of suggested MISI values','','\it Support \rm','---------------------------------------------------------------------------','If you are having any issues please let us know.','email us at shauna.odonovan@maastrichtuniversity.nl','','\it Citation \rm','--------------------------------------------------------------------------','\it O\textsc{\char13}Donovan et al. Improved quantification of muscle insulin sensitivity using oral glucose tolerance test data: the MISI Calculator. Scientific Reports (2019) 9:9388. \rm','---------------------------------------------------------------------------''This work was supported by the Dutch Province of Limburg','---------------------------------------------------------------------------','Copyright held by the Maastricht Centre for Systems Biology','This program is free software: you can redistribute it and/or modify it under the terms of the GNU Genral Public License as published by the Free Software Foundata, either version 3 of the Licence,or (at your option) and later version','This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY: without even the implied warranty of MERCHANTABILITY or FITTNESS FOR A PARTICULAR PURPOSE. See the GNU General Public Licence for more details','You should have received a copu of the GNU Genral Public Licence along with this program. If not, see<http://www.gnu.org/licences/>.'},'About',options);

% --- Executes on button press in plot_question.
function plot_question_Callback(hObject, eventdata, handles)
%--------------------------------------------------------------------------
%displays message on selection of help button/question button in step 6
%(plot flagged curves)
%--------------------------------------------------------------------------
options.Interpreter = 'tex';
options.WindowStyle='modal';

msgbox({'\bf Plotting flagged glucose curves \rm','','It is possible to visualise flagged glucose curve in the MISI Calculator tool by selecting the \it Plot \rm button. The glucose and insulin curves for flagged individuals will be displayed in the panels on the calculator interface along with graphical representations of the suggested MISI score values (the line defined by dG/dt. The title of the plot will contain the user supplied identifier (ID number), the reason for flagging, and the method by which MISI was computed. The user may choose to save a plot by selecting the \it save plot \rm button located on the upper right corner of the glucose plot. The figure will be saved to the user specified output directory. To plot the next flagged individual select the \it Plot \rm button again.'},'Help',options);
