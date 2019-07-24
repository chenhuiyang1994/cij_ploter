# Reading an excel file using Python 
import xlrd 
  
# Give the location of the file 
loc1 = ("stifness non relax.xlsx")
loc2 = ("stifness relax.xlsx")
loc3 = ("stifness phonon - Copy.xlsx")
loc4 = ("exp_data.xlsx")

# To open Workbook
wb1 = xlrd.open_workbook(loc1) 
sheet1 = wb1.sheet_by_index(0)

wb2 = xlrd.open_workbook(loc2) 
sheet2 = wb2.sheet_by_index(0)

wb3 = xlrd.open_workbook(loc3) 
sheet3 = wb3.sheet_by_index(0)
#print (sheet1.col(0),"sheet1")

wb4 = xlrd.open_workbook(loc4)
sheet4 = wb4.sheet_by_index(0)
#=====append exp data=========
exp_x=[]
exp_y=[]
for e in range (0,sheet4.nrows-1):
    exp_x.append(sheet4.cell_value(e,0))
    exp_y.append(sheet4.cell_value(e,1))
    
#==========typed in pressure==========
p=[0.049626103,
20.16511947,
40.30190761,
   50,
80.27413441,
   100,
120.1735292,
140.2123701
   #150
]
pp=p
#===========plots============
#=======
for k in range (0,9):
    cn=[]
    cr=[]
    cp=[]
#=======
    for i in range(1,sheet1.nrows-1): 
        cn.append(sheet1.cell_value(i, k))

    for i in range(1,sheet2.nrows-1): 
        cr.append(sheet2.cell_value(i, k))

    for i in range(1,sheet3.nrows-1): 
        cp.append(sheet3.cell_value(i, k))

    print(cn,cr,cp)

    import matplotlib.pyplot as plt

    #scatter
    plt.plot(pp,cn,'ro',label='non relax')
    plt.plot(p,cr,'bo',label='relax')
    plt.plot(pp,cp,'go',label='phonon')

    #exp fit
    def exponenial_func(t, c0, c1, c2,c3):
        return c0+c1*t-c2*np.exp(-c3*t)

    import numpy as np
    from scipy.optimize import curve_fit

    p0=[100,0.01,100,0.01]

    popt1, pcov = curve_fit(exponenial_func, pp, cn, p0,maxfev=100000)
    popt2, pcov = curve_fit(exponenial_func, p, cr, p0,maxfev=100000)
    popt3, pcov = curve_fit(exponenial_func, pp, cp, p0,maxfev=100000)

    xx = np.linspace(0, 150, 100)
    yy1 = exponenial_func(xx, *popt1)
    yy2 = exponenial_func(xx, *popt2)
    yy3 = exponenial_func(xx, *popt3)

    plt.plot(xx, yy1,'r')
    plt.plot(xx, yy2,'b')
    plt.plot(xx, yy3,'g')
    ax = plt.gca()
    plt.ylim(0,1400)
    plt.plot(exp_x,exp_y,'p')

    if k==0:
        plt.legend(loc='upper left')
        plt.title('comparison of c11')
        plt.savefig('comparison of c11.png')
        plt.show()
        plt.ylim(0,1800)

    elif k==1:
        plt.legend(loc='upper left')
        plt.title('comparison of c22')
        plt.savefig('comparison of c22.png')
        plt.show()
        plt.ylim(0,1800)
    
    elif k==2:
        plt.legend(loc='upper left')
        plt.title('comparison of c33')
        plt.savefig('comparison of c33.png')
        plt.show()
        plt.ylim(0,1800)

    elif k==3:
        plt.legend(loc='upper left')
        plt.title('comparison of c44')
        plt.savefig('comparison of c44.png')
        plt.show()
        plt.ylim(0,800)
        
    
    elif k==4:
        plt.legend(loc='upper left')
        plt.title('comparison of c55')
        plt.savefig('comparison of c55.png')
        plt.show()
        plt.ylim(0,800)
    


    elif k==5:
        plt.legend(loc='upper left')
        plt.title('comparison of c66')
        plt.savefig('comparison of c66.png')
        plt.show()
        plt.ylim(0,800)
    
    
    elif k==6:
        plt.legend(loc='upper left')
        plt.title('comparison of c12')
        plt.savefig('comparison of c12.png')
        plt.show()
        plt.ylim(0,800)

    elif k==7:
        plt.legend(loc='upper left')
        plt.title('comparison of c23')
        plt.savefig('comparison of c23.png')
        plt.show()
        plt.ylim(0,800)
    
    elif k==8:
        plt.legend(loc='upper left')
        plt.title('comparison of c13')
        plt.savefig('comparison of c13.png')
        plt.show()
        plt.ylim(0,800)

    else:
        pass
