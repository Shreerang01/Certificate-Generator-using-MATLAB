
clc       % Clear command window.
clear all % Clear variables and functions from memory
close all % closes all open windows
home      % Send the cursor home

% Read Excel sheet containing details. Text is read from the file
filename = 'Sample_list.xls';
[num,txt] = xlsread(filename) % seperately from numbers

% obtain length of text in excel or number of certificates to be generated
len=length(txt)

% Read blank certificate image
blankimage = imread('sample_certificate.tiff');


% Obtain names from the txt variable which are in 2nd column
for i=1:len
    for j= 2:2 
        text_names(i,j)=txt(i,j)
    end
end

% Obtain topics from the txt variable which are in 3rd column
for i=1:len
    for j= 3:3
      text_topic(i,j)=txt(i,j)
    end
end


%Ignore first row which is heading
for i=2:len
        text_all=[text_names(i,2) text_topic(i,3)] % combine names and topics
        
        % obtain positions to insert on image, MSPaint or any image editor
        position = [1035 703;1027 937];         
        
        %Provide parameters for function insertText
        RGB = insertText(blankimage,position,text_all,'Font','Times New Roman','FontSize',60,'BoxOpacity',0,TextColor="green");
        
        %Obtain and display figure in color
        figure
        imshow(RGB)        
        
        % generate and save files with .tif extension
        y=i-1
        filename=['Test_certificate' num2str(y)]
        lastf=[filename '.tif']
        saveas(gcf,lastf)
end    
