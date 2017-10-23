% ----------------------------------------------------------------------
% Digital Signal and Audio Processing
% Final Project: Acoustic Guitar-Pitch Detection 
%                using Cepstral Analysis Excel-GUI Analyzer
% 
% Group No. 2
% Leader:   Prince Julius T. Hari
% Members:  Lester Y. Besabe
%           Christian Erick L. Cimbracruz
%           Vanessa E. Espino
%           John Denver P. Felarca
%           Hannah Lalaine Morasa1
%           John Daryll M. Moya
%
% BSCS 4-4 AY 2017-18
% ----------------------------------------------------------------------

function export_data() 

    display(sprintf('\n\tChoose where to get the chord:\n'));
    display(sprintf('\t\t[ 0 ] Actual Record'));
    display(sprintf('\t\t[ 1 ] Search from files\n'));
    
    % Prompts for user choice
    prompt = sprintf('\tChoice: ');
    try
        choice = input(prompt);
    catch
        display(sprintf('\tInvalid input\n'));
    end % try
    
    if choice 
        get_from_files();
    else
        actual_record();
    end % if choice

end % function export_data()

function actual_record()

    SAMPLE_RATE = 44100; % Sample rate (Fs) in Hz
    SAMPLE_SIZE = 16;    % Sample size nBits
    NUM_CHANNEL = 1;     % Number of channels
    RECORDING_TIME = 5;
    
    rec_obj = audiorecorder(SAMPLE_RATE, SAMPLE_SIZE, NUM_CHANNEL);
    display(sprintf('\tStart recording.'));
    recordblocking(rec_obj, RECORDING_TIME);
    display(sprintf('\tEnd of recording.\n'));
    
    record_data = getaudiodata(rec_obj);
    filename = 'record.wav';
    audiowrite(filename, record_data, SAMPLE_RATE);
    
    % Read and export the chord/sound
    read_chord(filename);

end % actual_record()

function get_from_files()

    % Look for an audio file (chord's file)
    [file, path] = uigetfile('*.wav', 'Select chord (wave) file'); 
    [~, ~, ext] = fileparts(file);
    
    % Check if theres an audio file
    if isequal(file,0)
        display(sprintf('\n\tNo audio file found.'));
        display(sprintf('\tProgram is terminated.\n'));
    elseif isequal(ext,'.wav')
        chord = strcat(path,'\',file);
        % Read and export the chord/sound
        read_chord(chord);
    else
        display(sprintf('\n\tNot an .wav file.'));
        display(sprintf('\tProgram is terminated.\n'));
    end

end % get_from_files()

function read_chord(chord)

    XLFILE = 'guitar_pitch.xlsm'; % File where the data will be exported
    SHEET = 'Sheet1';             % Sheet name of the file
    RANGE1 = 'B3';                % Start cell where data will be written
    RANGE2 = 'E3';                % Cell where sample rate will be written
    RANGE3 = 'E4';                % Cell where no. samples will be written
    no_error = 1;                 % Flag if theres no error
        
    % Read audio file
    [y, Fs] = audioread(chord);
    display(y); 
    plot(y);
    
    try
        xlswrite(XLFILE, y, SHEET, RANGE1);         % Write raw data
        xlswrite(XLFILE, Fs, SHEET, RANGE2);        % Write sampling Rate
        xlswrite(XLFILE, length(y), SHEET, RANGE3); % NWrite no. of samples
    catch
        display(sprintf('\tA problem encountered while exporting the'));
        display(sprintf('\traw data to %s\n', XLFILE));
        no_error = 0;
    end
    
    if no_error
        display(sprintf('\tSuccessfully exported the'));
        display(sprintf('\traw data to %s\n', XLFILE));
    end
        
end % function record_chord()