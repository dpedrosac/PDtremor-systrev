n = 6;
no_tot = [5137];
nos = [1250, 3887, 185, 0, 0, 0];

%% General settings
tot_h = 500; tot_w = 8;
width = 2; height = 55;
posleft = bsxfun(@times, repmat([1 15 width height], [n,1]), [1, NaN, 1, 1]);
posleft(:,2) = linspace(.8*tot_h, .05*tot_h, n);

%% Check numbers and definen the categories
% if nansum(nos(1:end-1))~=no_tot(1)-no_tot(2)
%     error('the numbers provided is not correct as the sum is not identical')
% else 
    nos = [nos, nansum(nos(1:end-1))];
%end

txt = {sprintf('PsycINFO\n(n=%d)', nos(1)), ...
    sprintf('MEDLINE\n(n=%d)', nos(2)), ...
    sprintf('All references \n(n=%d)', nansum(nos(1:2))-nos(3)), ...
    '', ...
    '', ...
    ''};

txt_right = {sprintf('Duplicates \n(n=%d)', nos(3)), ...
    '', ...
    '', ...
    '', ...
    '', ...
    ''};

%% Define the positions of both columns
posright = bsxfun(@times, repmat([3.5 15 width height], [n-1,1]), [1, NaN, 1, 1]);
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
    if k>1
        rectangle('Position', posright(k-1,:));
        text(posright(k-1,1)+width/2, posleft(k-1,2)-height/4, txt_right{k-1}, 'HorizontalAlignment', 'Center')
    end
    text(posleft(k,1)+width/2, posleft(k,2)+height/2, txt{k}, 'HorizontalAlignment', 'Center')
end
axis([0 tot_w 0 tot_h])

% Start plotting arrows
for k = 1:n-1
    line([1 1].*posleft(1,1)+.5*width, [posleft(k,2) posleft(k+1,2)+height], 'Color', 'k');
    line([posleft(k,1)+.5*width, posright(k,1)], [1 1].*posright(k,2)+height/2, 'Color', 'k');
end
