clc;clear all;

function [aRMS_X, aRMS_Y, aRMS_Z, aRMS_SUM] = req_aRMS(filename, sheetname)
data = xlsread(filename, sheetname);
f   = data(2:length(data(:,1)),1); %Frequency
p_x = data(2:length(data(:,1)),2); %PSD function
p_y = data(2:length(data(:,1)),3);
p_z = data(2:length(data(:,1)),4);
aRMS_X = cal_aRMS_horizontal(f, p_x);
aRMS_Y = cal_aRMS_horizontal(f, p_y);
aRMS_Z = cal_aRMS_vertical(f, p_z);
aRMS_SUM = sqrt((1.4*aRMS_X)^2 + (1.4*aRMS_Y)^2 + aRMS_Z^2);
fprintf('%-40s |    %-8.3f | \n', 'X-axis RMS weighted acceleration',     aRMS_X);
fprintf('%-40s |    %-8.3f | \n', 'Y-axis RMS weighted acceleration',     aRMS_Y);
fprintf('%-40s |    %-8.3f | \n', 'Z-axis RMS weighted acceleration',     aRMS_Z);
fprintf('%-40s |    %-8.3f | \n', 'Sum of all RMS weighted acceleration', aRMS_SUM);
end

%% Frequency Weighting Function
function [aRMS] = cal_aRMS_vertical(f, p) %Weighted RMS Acceleration for vertical
    ww = zeros;
    acc_num = 0;
    for i = 1:length(f) 
        if f(i) >= 0.5 && f(i) < 2
            ww(i) = 0.5;
        elseif f(i) >= 2 && f(i) < 4
            ww(i) = f(i) / 4;
        elseif f(i) >= 4 && f(i) < 12.5
            ww(i) = 1;
        elseif f(i) >= 12.5 && f(i) < 80
            ww(i) = 12.5 / f(i);
        else
            ww(i) = 0;
        end
        acc_num = acc_num + ww(i)^2 * p(i) * (f(2) - f(1));
    end
    aRMS = (acc_num)^0.5;
end

function [aRMS] = cal_aRMS_horizontal(f, p) %Weighted RMS Acceleration for horizontal
    ww = zeros;
    acc_num = 0;
    for i = 1:length(f) 
        if f(i) >= 0.5 && f(i) < 2
            ww(i) = 1;
        elseif f(i) >= 2 && f(i) < 80
            ww(i) = 2 / f(i);
        else
            ww(i) = 0;
        end
        acc_num = acc_num + ww(i)^2 * p(i) * (f(2) - f(1));
    end
    aRMS = (acc_num)^0.5;
end