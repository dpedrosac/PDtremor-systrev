n           = 5;                                                            % levels to use
nos         = [1336, 5823, 4119, 3385, 7540, 31, 0];                      
nos_man     = [0, 0, 0, 0, 19, 0, 188];                                       % numbers of manually added studies
excl        = [59, 58, 20, 22, 11, 9, 15, 6, 1, 12, 0, 4];
excl_man    = [5, 87, 0, 26, 3, 0, 0, 2, 0, 1, 39, 0];

add_manual  = 1;
switch add_manual; case(1); nos = nos + nos_man; excl = excl + excl_man; end

%% General settings
tot_h = 500; tot_w = 8;
width = 2; height = 55;
posleft = bsxfun(@times, repmat([1 15 width height], [n,1]), [1, NaN, 1, 1]);
posleft(:,2) = linspace(.8*tot_h, .05*tot_h, n);
ftsize = 16;
%% Check numbers and definen the categories
% if nansum(nos(1:end-1))~=no_tot(1)-no_tot(2)
%     error('the numbers provided is not correct as the sum is not identical')
% else 
    nos = [nos, nansum(nos(1:end-1))];
%end

txt = {sprintf('PsycINFO\n(n=%d)\nMEDLINE\n(n=%d)\nEmbase\n(n=%d)', nos(1), nos(3), nos(2)), ...
    sprintf('All references \n(n=%d)', nansum(nos([1:3, 7]))), ...
    sprintf('Processed \n(n=%d)', nansum(nos([1:3, 7])) - nos(4)), ...
    sprintf('Full-text search \n(n=%d)', nansum(nos([1:3, 7])) - nansum(nos(4:6))), ...
    sprintf('Processed \n(n=%d)', nansum(nos([1:3, 7])) - nansum(nos(4:6)) - nansum(excl))};

txt_right = {'', ...
    sprintf('Duplicates \n(n=%d)', nos(4)), ...
    sprintf('Excluded after abstract\nscreening (n=%d)\n\nExcluded after \n discussion (n=%d)', nos(5), nos(6)), ...
    sprintf('%d Excluded because:\nReview (n=%d)\nNo tremor outcome (n=%d)\nNot enough \nparticipants(n=%d)\nStudy not found (n=%d)\nCase report(n=%d)\nWrong intervention (n=%d)\nNo intervention \noccured (n=%d)\nDuplicate study (n=%d)\nWrong population \nstudied (n=%d)\nIncomplete results \nprovided (n=%d)\nPosters/Abstracts(n=%d)\nOther reasons (n=%d)', ...
    nansum(excl), excl(1), excl(2), excl(3), excl(4), excl(5), ...
    excl(6), excl(7), excl(8), excl(9), excl(10), ...
    excl(11), excl(12))};

%% Define the positions of both columns
posright = bsxfun(@times, repmat([3.5 15 width height], [n-1,1]), [1, NaN, 1.3, 1]);
iter = 0;
for k = 1:n-1
    iter = iter+1;
    posright(k,2) = posleft(k,2) - abs(diff(posleft(1:2,2))).*.5;
end

%% Start plotting all data
% Start plotting rectangles
figure(1); clf; hold on;
axis off
for k = 1:n
    rectangle('Position', posleft(k,:));
    if k == 5
        rectangle('Position', posright(k-1,:).*[1,0.02,1,2.5]);
        text(posright(k-1,1)+width/8, posleft(k-1,2)-height/1.1, ...
            txt_right{k-1}, 'HorizontalAlignment', 'Left', 'FontSize', 14)
    elseif k > 2
        rectangle('Position', posright(k-1,:));
        text(posright(k-1,1)+width/8, posleft(k-1,2)-height/4, ...
            txt_right{k-1}, 'HorizontalAlignment', 'Left', 'FontSize', 14)
    end
    text(posleft(k,1)+width/2, posleft(k,2)+height/2, txt{k}, ...
        'HorizontalAlignment', 'Center', 'FontSize', 14)
end
axis([0 tot_w 0 tot_h])

% Start plotting arrows
for k = 1:n-1
    line([1 1].*posleft(1,1)+.5*width, [posleft(k,2) posleft(k+1,2)+height], 'Color', 'k');
    line([posleft(k,1)+.5*width, posright(k,1)], [1 1].*posright(k,2)+height/2, 'Color', 'k');
end

% Add two more rectangles on top:
rectangle('Position', [3.5, 400, 2, 55])
text(3.5+width/2, 400+height/2, sprintf('Manually\n added(n=%d)', nos(7)), ...
    'HorizontalAlignment', 'Center', 'FontSize', 14)
rectangle('Position', [0.1, 305, .5, 150])
text(.35, 380, 'Identification', 'HorizontalAlignment', 'Center', 'Rotation', 90, ...
    'FontWeight', 'bold', 'FontSize', ftsize)
rectangle('Position', [0.1, 110, .5, 150])
text(.35, 185, 'Screening', 'HorizontalAlignment', 'Center', 'Rotation', 90, ...
    'FontWeight', 'bold', 'FontSize', ftsize)
rectangle('Position', [0.1, 10, .5, 75])
text(.35, 50, 'Processing', 'HorizontalAlignment', 'Center', 'Rotation', 90, ...
    'FontWeight', 'bold', 'FontSize', ftsize)

