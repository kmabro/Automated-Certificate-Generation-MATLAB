%Khan Muhammad (KM ABRO) 

clc;
clear;
close all;

blankimage = imread('Certificate_Blank.png');

file = 'Data.xlsx';
[num, txt] = xlsread(file);

len = size(txt, 1); 

for i = 2:len
    name = txt{i, 2};   
    topic = txt{i, 3};
    
    base_x = 960; 
    text_width = length(name) * 22; 
    name_x = base_x - text_width / 2;

    name_position = [960, 750]; 
    topic_position = [1400, 840]; 
    
    RGB = insertText(blankimage, name_position, name, ...
        'FontSize', 90, 'TextColor', 'white', 'BoxOpacity', 0, 'AnchorPoint', 'Center');
    
    RGB = insertText(RGB, topic_position, topic, ...
        'FontSize', 35, 'TextColor', 'white', 'BoxOpacity', 0, 'AnchorPoint', 'Center');
    
    imshow(RGB);
    
   
    filename = ['Certificate_ ' strrep(name, ' ', ' ') '.png'];
    imwrite(RGB, filename);
end
%%