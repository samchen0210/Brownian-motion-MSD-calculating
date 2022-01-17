clear all; close all;
clc;
%%%%%%%%%%%%%%% Program editior : Sam Chen, 2022/01/17 %%%%%%%%%%%%%%% 
%%%%%%%%%%%%%%% Contact: B812107036@tmu.edu.tw%%%%%%%%%%%%%%%
%% Parameter setting, START   
aa = 'Define interval numbers? ';
x = input(aa);
filename = 'FINAL RESULT.xlsx'; [status,sheets] = xlsfinfo(filename);
sheet_num = size(sheets,2); %sheets = {'FREE','LASER','GROUP1-1'...};
signal_num = 200;  %The length of the figure
points = x; %interval set 

%% Import data of particle tracking, START
xdata= zeros(sheet_num,  signal_num);  ydata= zeros(sheet_num, signal_num);
for sheet_index = 1: sheet_num
    data = readmatrix(filename,'sheet', string(sheets(sheet_index)), 'Range', 'B3:C202');
    sheet_name = string(sheets(sheet_index));
    for index =1:signal_num
        %sheet(3) = 'CONTROL'%run analysising and plotting the result
        xdata(sheet_index,index) = data(index, 1)';
        ydata(sheet_index,index) = data(index ,2)';
    end
    MSD(sheets, sheet_index, sheet_name, signal_num, points,xdata,ydata) 
end

 %% calculating MSD,START
 %%%% Experiment steps in followin scripts : %%%%
 %%%%  1. Compute different time intervals   %%%%
 %%%%  2. Average the data respectively          %%%%  
 %%%%  3. Plot the data and show figure           %%%%
function  MSD(sheets, sheet_index, sheet_name, signal_num, points,xdata,ydata)  
sheet_num = size(sheets,2);    
X= zeros(points, signal_num,sheet_num);  Y= zeros(points, signal_num,sheet_num);
    for interval = 1:points
        for index = 1:(signal_num-1)
            if  (interval+index )<= signal_num
                 X(interval, index,sheet_index) = (xdata(sheet_index, index+interval) - xdata(sheet_index , index))^2;
                 Y(interval, index,sheet_index) = (ydata(sheet_index, index+interval) - ydata(sheet_index , index))^2;
             end 
         end
    end

    X_average = zeros(points ,1,sheet_index);   Y_average = zeros(points ,1 ,sheet_index); 
    dr = zeros(points, 1,  sheet_index);
    for interval =1: points
        for index = 1: (signal_num-1)
            if (X(interval,index,sheet_index) ~=0) && (Y(interval,index,sheet_index) ~= 0)
                X_average(interval,sheet_index)=mean(X(interval,:,sheet_index));
                Y_average(interval,sheet_index)=mean(Y(interval,:,sheet_index));
                dr(interval, 1,sheet_index) = (X_average(interval, : ,sheet_index))+(Y_average(interval ,: ,sheet_index));
            end
         end
    end
%%%%%%%%%% calculating MSD,END%%%%%%%%%%

%%%%%%%%%%%%%%ploting the data,Start%%%%%%%%%%%%%%
    figure
    x_interval = linspace(1,points,points)'; 
    y_interval = dr(: , sheet_index)
    plot(x_interval,y_interval,'-r','LineWidth',4);
    graph_name = append(sheet_name, ' (', num2str(points), '-intervals) ')
    title(strcat("Brownian motion experiment : ",graph_name));
    xlabel('Time interval'); ylabel('MSD (x nm^2)');
    %set(gcf,'color','none');%set(gca,'color','none');
    grid on; box on;
%%%%%%%%%%%%%%ploting the data,END%%%%%%%%%%%%%%

end

%%%%%%%%%%%%%%% Program editior : Sam Chen, 2022/01/17 %%%%%%%%%%%%%%% 
%%%%%%%%%%%%%%% Contact: B812107036@tmu.edu.tw%%%%%%%%%%%%%%%
