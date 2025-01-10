% MATLAB Script for Processing Seed Volume and Damage Data
% Author: Sherif HAMDY
% Date: 28/09/2020
% Description: This script reads input data, processes seed volume and damage information,
% calculates statistics, and writes results into Excel files.

clear; close all; clc;

% === Input Setup ===
inputFolder = 'C:\PathToExcelFiles';
outputFileAllData = 'Results_Data.xlsx';
outputFileSummary = 'Results_Summary.xlsx';

% Load input data
TRTAB = readtable('C:\PathToLotsOrder.xlsx');
TRTAB_C = table2cell(TRTAB);
SAMPLELIST = unique(cell2mat(TRTAB_C(~cellfun(@isempty, TRTAB_C(:,1)), 1)));

% Headers for output data
headerAllData = {'Lot', 'VolumeSeed', 'VolumeDamage'};
headerSummary = {'Lot', 'NumSeeds', 'MeanSeedVol', 'SDSeedVol', ...
                 'MeanUndamagedVol', 'SDUndamagedVol', 'MeanDamagedVol', ...
                 'SDDamagedVol', 'MeanCorrectedDamagedVol', ...
                 'SDCorrectedDamagedVol', '%DamagedSeeds', ...
                 '%SeedDamage', '%LotDamage'};

% Initialize data storage
allData = headerAllData;
summaryData = headerSummary;

% Change to input folder
cd(inputFolder);

% Get list of files
fileList = dir('*S.csv');

% === Main Processing Loop ===
for i = 1:length(fileList)
    % Construct file paths
    principalFile = sprintf('SCAN_400%d_%d_data_S.csv', TRTAB_C{i,1}, TRTAB_C{i,2});
    associatedFile = sprintf('SCAN_400%d_%d_data_D.csv', TRTAB_C{i,1}, TRTAB_C{i,2});

    % Read data from CSV files
    volumePrincipal = csvread(principalFile, 2, 0); % Skip 2 header rows
    volumeAssociated = csvread(associatedFile, 2, 0);

    % Ensure equal lengths (truncate if necessary)
    n = min(length(volumePrincipal), length(volumeAssociated));
    volumePrincipal = volumePrincipal(1:n);
    volumeAssociated = volumeAssociated(1:n);

    % Combine data
    data = [repmat(TRTAB_C{i,2}, n, 1), volumePrincipal, volumeAssociated];

    % Filter data based on thresholds
    data(data(:, 2) < 30 | data(:, 2) > 800, :) = [];

    % Store cleaned data
    allData = [allData; num2cell(data)];

    % Statistical calculations
    numSeeds = size(data, 1);
    meanSeedVol = mean(data(:, 2));
    sdSeedVol = std(data(:, 2));
    meanUndamagedVol = mean(data(data(:, 3) < 10, 2));
    sdUndamagedVol = std(data(data(:, 3) < 10, 2));
    meanDamagedVol = mean(data(data(:, 3) >= 10, 2));
    sdDamagedVol = std(data(data(:, 3) >= 10, 2));
    meanCorrectedDamagedVol = mean(data(data(:, 3) >= 10, 2) + data(data(:, 3) >= 10, 3));
    sdCorrectedDamagedVol = std(data(data(:, 3) >= 10, 2) + data(data(:, 3) >= 10, 3));
    pctDamagedSeeds = (sum(data(:, 3) >= 10) / numSeeds) * 100;
    pctSeedDamage = mean(data(data(:, 3) >= 10, 3) ./ ...
                        (data(data(:, 3) >= 10, 2) + data(data(:, 3) >= 10, 3))) * 100;
    pctLotDamage = (sum(data(data(:, 3) >= 10, 3)) / ...
                   (sum(data(:, 2)) + sum(data(data(:, 3) >= 10, 3)))) * 100;

    % Append summary statistics
    summaryData = [summaryData; ...
                   {TRTAB_C{i,2}, numSeeds, meanSeedVol, sdSeedVol, ...
                    meanUndamagedVol, sdUndamagedVol, meanDamagedVol, ...
                    sdDamagedVol, meanCorrectedDamagedVol, ...
                    sdCorrectedDamagedVol, pctDamagedSeeds, ...
                    pctSeedDamage, pctLotDamage}];
end

% === Write Output ===
xlswrite(outputFileAllData, allData);
xlswrite(outputFileSummary, summaryData);

% === Cleanup ===
clear; clc;
disp('Processing complete. Results saved.');
