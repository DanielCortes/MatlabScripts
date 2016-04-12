function [ vectors ] = VecGen( strain )
%VecGen generates vectors from split tracks to make Quiver vectors
%   DGenerates x,y,u,v vectors for Quiver from U-track matrices
Mat = xlsread('Go5_1_245b.xls');
SizeVec = size(Mat);
Size = SizeVec(1);
RSize = (Size / 8 ) - 1
po = 0;
n = 0;
Xi = 0
Xf = 0
Yi = 0
Yf = 0
for i = 1:RSize
    n = 8 * po + 1;
    m = 8 * po + 2;
    Xi = [Xi;Mat(n)]
    Xf = [Xf;Mat(n,2)]
    Yi = [Yi;Mat(m)]
    Yf = [Yf;Mat(m,2)]
    po = po + 1;
    prog = (po / RSize) * 100;
    fprintf('Progress: %5.2f\n', prog)
    my_nameXi = sprintf('Xi%s', strain);
    my_nameXf = sprintf('Xf%s', strain);
    my_nameYi = sprintf('Yi%s', strain);
    my_nameYf = sprintf('Yf%s', strain);
    xlswrite(my_nameXi, Xi)
    xlswrite(my_nameXf, Xf)
    xlswrite(my_nameYi, Yi)
    xlswrite(my_nameYf, Yf)
end

