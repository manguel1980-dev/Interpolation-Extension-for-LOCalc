# Created by: Manuel V. Astros M. (manuel.astros1980@gmail.com) July 2018.

# The objective of this add-in is to create an interpolations functions.
# All functions are developed to behave in the same way that the interpolation add-in created by Jose Ivan Martinez Garcia for Excel.
# It will make possible a kind of interoperability among LibreOffice-Calc, Google-sheet in interpolation matter.

# The Excel Interpolation Addin developed by Jose Ivan Martinez Garcia (martinji@unican.es)
# can be found published in http://personales.gestion.unican.es/martinji/Interpolation.htm

import uno
import unohelper
from Astros.Montezuma.CalcTools.CalcFunctions import XInterpolation
from math import fmod


class InterpolationImpl( unohelper.Base, XInterpolation ):
    def __init__(self, ctx):
        self.ctx = ctx

    # Function to linearly interpolate and extrapolate in base ordered from a domain and range arrangement
    def interpo(self, x, lRangeX, lRangeY):
        Leng = len(lRangeX) - 1     #get the total length of the domine arrangement less one
        if  Leng >= 1:
            dif = lRangeX[0][0] - lRangeX[1][0]     #It identify if range is ascendant or descendant
            #The way to work the array lRange[Num Row][Num column]
            if dif == 0:
                Interp = "Elements in domine range can not be equals"
                #msgbox("Elements in domine range can not be equals")
                return Interp
            for k in range(1, Leng):        #Test that the array is monotonic
                if Leng == 1:
                    break       # no need use "break", due to arrangement only contains two elements
                difn = lRangeX[k][0] - lRangeX[k + 1][0]        #Here prove that the rest element of the arrangement are the same sign of the first elements
                prox = dif * difn
                if prox <= 0:
                    Interp = "Domain must be monotonic"
                    #msgbox("Domain must be monotonic")
                    return Interp
            xp1 = -1         # Counter is inicialized

            for d in range(Leng):       #This loop search values that contain the argument or take two values at the beginning or at the end
                if dif < 0:
                    if lRangeX[d][0] > x:
                        break
                else:
                    if lRangeX[d][0] < x:
                        break
                xp1 = xp1 + 1

            if xp1 >= Leng:
                xp1 = Leng - 1

            if xp1 < 1:
                xp1 = 0

            x0 = lRangeX[xp1][0]
            x1 = lRangeX[xp1+1][0]
            y0 = lRangeY[xp1][0]
            y1 = lRangeY[xp1+1][0]
            Interp = y0 + ((x - x0) * (y1 - y0) / (x1 - x0))          #Calculate linear interpolation
        else:
            Interp = "Only one item in the domine"
            #msgbox("Only one item in the domine")
        return Interp
    #######################- END METHOD -############################



    # Function to linearly interpolate considering one array table that contains the domain and ranges to interpolate and the two entrances.
    #The way to work the array lRange[Num Row][Num column]
    def interpo2(self, x, y, lRange):
        LenR = len(lRange) - 1          #Number of rows
        LenC = len(lRange[0]) - 1       #Number of columns

        if LenR < 2 or LenC < 2:        #Detect if the number of rows and columns are higher than 1
            Interp2 = "Not enough rows or column"
            #msgbox("Not enough rows or column")
            return Interp2
        #Rows are going to be check first
        diffR = lRange[1][0] - lRange[2][0]     #it help to detect whether is Ascendant or Descendant
        for k in range(1, LenR):
            if LenR == 2:       #if the function only have two rows. No need continue with "for" statement
                break
            diffN = lRange[k][0] - lRange[k + 1][0]        #Here prove that the rest element of the arrangement are the same sign of the first elements
            prox = diffR * diffN
            if prox <= 0:
                Interp2 = "Rows in function argument shall be monotonic"
                #msgbox("Rows in function argument shall be monotonic")
                return Interp2

        xp1 = -1
        for d in range(1, LenR):        #This loop search values that contain the argument or take two values at the beginning or at the end
            if diffR < 0:
                if lRange[d][0] > x:
                    break
            else:
                if lRange[d][0] < x:
                    break
            xp1 = xp1 + 1

        if xp1 > LenR:
            xp1 = LenR

        if xp1 < 0:
            xp1 = 0

        xp1 = xp1 + 1
        #Interp2 = xp1

        #Columns are going to be checked first
        diffC = lRange[0][1]-lRange[0][2]       #it help to detect whether is Ascendant or Descendant
        for k in range(1, LenC):
            if LenC == 1:       #function is monotonic, only contain two (2) elements
                break
            diffN = lRange[0][k] - lRange[0][k + 1]        #Here prove that the rest element of the arrangement are the same sign of the first elements
            prox = diffC * diffN
            if prox <= 0:
                Interp2 = "Column in function argument shall be monotonic"
                #msgbox("Column in function argument shall be monotonic")
                return Interp2

        yp1 = -1
        for d in range(1, LenC):        #This loop search values that contain the argument or take two values at the beginning or at the end
            if diffC < 0:
                if lRange[0][d] > y:
                    break
            else:
                if lRange[0][d] < y:
                    break
            yp1 = yp1 + 1

        if yp1 > LenC:
            yp1 = LenC

        if yp1 < 0:
            yp1 = 0

        yp1 = yp1 + 1

        x0 = lRange[xp1][0]
        x1 = lRange[xp1 + 1][0]
        y0 = lRange[0][yp1]
        y1 = lRange[0][yp1 + 1]
        a = lRange[xp1][yp1]
        b = lRange[xp1][yp1 + 1]
        chge = lRange[xp1 + 1][yp1]
        d = lRange[xp1 + 1][yp1 + 1]
        p1 = a + (y - y0) / (y1 - y0) * (b - a)
        p2 = chge + (y - y0) / (y1 - y0) * (d - chge)
        Interp2 = p1 + (x - x0) / (x1 - x0) * (p2 - p1)
        return Interp2
    #######################- END METHOD -############################


    #Function is used for interpolation or extrapolation using cubic polynomial functions, that adapt by pieces to the points
    #where it is necessary to interpolate. First derivative (slope) and second derivative at the ends of the polynomials match
    #with the next one and the values at the start of first and the end of the last splines can be setted on the basis of the
    #type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of Range_xy (1 column) and (2nd key and V2) for the greater value of Range_xy (1st column).
    def cercha(self, x, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerch = "Range selection is incorrect"
            return cerch

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerch = "key argument larger than 3 is not allowed"
            return cerch

        if type(v1) is str or type(v2) is str:
            cerch = "v1 and v2 values must be numeric values"
            return cerch

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerch = "Number of row is an even number"
                return cerch

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerch = "Data must be order in ascending"
                return cerch

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= x and j < N:
            j += 1

        if j > 0:
            j -= 1
        cerch = ((Coef[j][0] * (x - lRange[j][0]) + Coef[j][1]) * (x - lRange[j][0]) + Coef[j][2]) * (x - lRange[j][0]) + lRange[j][1]
        return cerch
    #######################- END METHOD -############################


    #Function to determine the slope (1st derivative) at the initial point of first spline - Matlab ® (function csape).

    #Suggestion: Matlab (function csape) uses, by default, to the slopes of the interpolation splines, which would have a spline
    #with only the first four given points (for the initial slope) and the last four (for the final). For a similar calculation,
    #this function can be used previously selecting a Range_xy with those 4 points and typing end-point constraints "ee" (Lagrange'sconditions).
    #It will assign only a cubic one for these 4 points and later using the function CERCHAPF with the 4 last given points, in an
    #similar way, in order to obtain the final slope.
    #Finally, with the calculated slopes, the function to be used is CERCHA with the constraints "ff" and
    #the values calculated for v1 and v2.
    def cerchapi(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchpi = "Range selection is incorrect"
            return cerchpi

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchpi = "key argument larger than 3 is not allowed"
            return cerchpi

        if type(v1) is str or type(v2) is str:
            cerchpi = "v1 and v2 values must be numeric values"
            return cerchpi

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchpi = "Number of row is an even number"
                return cerchpi

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchpi = "Data must be order in ascending"
                return cerchpi

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)
        cerchpi = Coef[0][2]

        return cerchpi
        #######################- END METHOD -############################


    #This function determine the final slope at the last point of the last spline.
    def cerchapf(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchpf = "Range selection is incorrect"
            return cerchpf

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchpf = "key argument larger than 3 is not allowed"
            return cerchpf

        if type(v1) is str or type(v2) is str:
            cerchpf = "v1 and v2 values must be numeric values"
            return cerchpf

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchpf = "Number of row is an even number"
                return cerchpf

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchpf = "Data must be order in ascending"
                return cerchpf

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)
        if key == False:
            cerchpf = c[N] * h[N - 1] / 3 + c[N - 1] * h[N - 1] / 6 + did[N - 1]
        else:
            cerchpf = b[LenR - 1][0] * h[N - 1] / 3 + b[N - 1][0] * h[N - 1] / 6 + did[N - 1]

        return cerchpf
        #######################- END METHOD -############################


    #This function determine the initial curvature (2nd derivative) at the first point of the first spline.
    def cerchaci(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchci = "Range selection is incorrect"
            return cerchci

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchci = "key argument larger than 3 is not allowed"
            return cerchci

        if type(v1) is str or type(v2) is str:
            cerchci = "v1 and v2 values must be numeric values"
            return cerchci

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchci = "Number of row is an even number"
                return cerchci

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchci = "Data must be order in ascending"
                return cerchci

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)
        if key == False:
            cerchci = c[0]
        else:
            cerchci = b[0][0]

        return cerchci
        #######################- END METHOD -############################


    #This function determine the final curvature (2nd derivative) at the last point of the last spline.
    def cerchacf(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchcf = "Range selection is incorrect"
            return cerchcf

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchcf = "key argument larger than 3 is not allowed"
            return cerchcf

        if type(v1) is str or type(v2) is str:
            cerchcf = "v1 and v2 values must be numeric values"
            return cerchcf

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchcf = "Number of row is an even number"
                return cerchcf

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchcf = "Data must be order in ascending"
                return cerchcf

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)
        if key == False:
            cerchcf = c[N]
        else:
            cerchcf = b[LenR - 1][0]

        return cerchcf
    #######################- END METHOD -############################

    ############################### SECOND REELASE #########################################
        #This function determine the curvature radius of the segments in the given points.
    def cerchara(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchra = tuple([["Range selection is incorrect"] for i in range(LenR)])
            return cerchra

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchra = tuple([["key argument larger than 3 is not allowed"] for i in range(LenR)])
            return cerchra

        if type(v1) is str or type(v2) is str:
            cerchra = tuple([["v1 and v2 values must be numeric values"] for i in range(LenR)])
            return cerchra

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchra = tuple([["Number of row is an even number"] for i in range(LenR)])
                return cerchra

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchra = tuple([["Data must be order in ascending"] for i in range(LenR)])
                return cerchra

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        pen = [0 for i in range(LenR)]                                      #inicializing pen variable
        ra = [[0] for i in range(LenR)]                                     #inicializing ra variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        if t1 == "P" and t2 == "G":
            cerchra = tuple([["Rectas?"] for i in range(LenR)])
            return cerchra

        for i in range(N):
            pen[i] = Coef[i][2]

        if key == False:
            pen[LenR - 1] = c[N] * h[N - 1] / 3 + c[N - 1] * h[N - 1] / 6 + did[N - 1]
            for i in range(LenR):
                try:
                    ra[i][0] = (pen[i]**2 + 1)**(1.5) / c[i]
                except ZeroDivisionError:
                    ra[i][0] = 'Infinite'
        else:
            pen[LenR - 1] = b[LenR - 1][0] * h[N - 1] / 3 + b[N - 1][0] * h[N - 1] / 6 + did[N - 1]
            for i in range(LenR):
                try:
                    ra[i][0] = (pen[i]**2 + 1)**(1.5) / b[i][0]
                except ZeroDivisionError:
                    ra[i][0] = 'Infinite'

        cerchra = tuple(ra)

        return cerchra
    #######################- END METHOD -############################


            #This function determine the centers of curvature coordinates of the segment, in the given points.
    def cercharaxy(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchraxy = tuple([["Range selection is incorrect"] for i in range(LenR)])
            return cerchraxy

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchraxy = tuple([["key argument larger than 3 is not allowed"] for i in range(LenR)])
            return cerchraxy

        if type(v1) is str or type(v2) is str:
            cerchraxy = tuple([["v1 and v2 values must be numeric values"] for i in range(LenR)])
            return cerchraxy

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchraxy = tuple([["Number of row is an even number"] for i in range(LenR)])
                return cerchraxy

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchraxy = tuple([["Data must be order in ascending"] for i in range(LenR)])
                return cerchraxy

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        pen = [0 for i in range(LenR)]                                      #inicializing pen variable
        raxy = [[0 for j in range(2)] for i in range(LenR)]                  #inicializing raxy variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        if t1 == "P" and t2 == "G":
            cerchraxy = tuple([["Rectas?"] for i in range(LenR)])
            return cerchraxy

        for i in range(N):
            pen[i] = Coef[i][2]

        if key == False:
            pen[LenR - 1] = c[N] * h[N - 1] / 3 + c[N - 1] * h[N - 1] / 6 + did[N - 1]
            for i in range(LenR):
                try:
                    raxy[i][0] = lRange[i][0] - ((pen[i]**2 + 1)*pen[i] / c[i])
                    raxy[i][1] = lRange[i][1] + ((pen[i]**2 + 1) / c[i])
                except ZeroDivisionError:
                    raxy[i][0] = 0
        else:
            pen[LenR - 1] = b[LenR - 1][0] * h[N - 1] / 3 + b[N - 1][0] * h[N - 1] / 6 + did[N - 1]
            for i in range(LenR):
                try:
                    raxy[i][0] = lRange[i][0] - ((pen[i]**2 + 1)*pen[i] / b[i][0])
                    raxy[i][1] = lRange[i][1] + ((pen[i]**2 + 1) / b[i][0])
                except ZeroDivisionError:
                    raxy[i][0] = 0

        cerchraxy = tuple(raxy)

        return cerchraxy
    #######################- END METHOD -############################


        #This function determine the slopes (1st derivative) at the given (well-known) points.
    def cerchap(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c                    # fixed 20181122
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchp = tuple([["Range selection is incorrect"] for i in range(LenR)])
            return cerchp

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchp = tuple([["key argument larger than 3 is not allowed"] for i in range(LenR)])
            return cerchp

        if type(v1) is str or type(v2) is str:
            cerchp = tuple([["v1 and v2 values must be numeric values"] for i in range(LenR)])
            return cerchp

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchp = tuple([["Number of row is an even number"] for i in range(LenR)])
                return cerchp

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchp = tuple([["Data must be order in ascending"] for i in range(LenR)])
                return cerchp

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        pen = [[0] for i in range(LenR)]                                      #inicializing pen variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        for i in range(N):
            pen[i][0] = Coef[i][2]

        if key == False:
            pen[LenR - 1][0] = c[N] * h[N - 1] / 3 + c[N - 1] * h[N - 1] / 6 + did[N - 1]
        else:
            pen[LenR - 1][0] = b[LenR - 1][0] * h[N - 1] / 3 + b[N - 1][0] * h[N - 1] / 6 + did[N - 1]

        cerchp = tuple(pen)

        return cerchp
    #######################- END METHOD -############################


    #This function determine the polynomial (spline) coeficient.
    #The range selected to get theresoult shall be 3 or 4 colums
    #and equal number of rows than the polynomials needed.
    def cerchac(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchc = tuple([["Range selection is incorrect" for j in range(4)] for i in range(LenR)])
            return cerchc

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchc = tuple([["key argument larger than 3 is not allowed" for j in range(4)] for i in range(LenR)])
            return cerchc

        if type(v1) is str or type(v2) is str:
            cerchc = tuple([["v1 and v2 values must be numeric values" for j in range(4)] for i in range(LenR)])
            return cerchc

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchc = tuple([["Number of row is an even number" for j in range(4)] for i in range(LenR)])
                return cerchc

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchc = tuple([["Data must be order in ascending" for j in range(4)] for i in range(LenR)])
                return cerchc

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        cerchc = tuple(Coef)

        return cerchc
    #######################- END METHOD -############################


    #This function determine the polynomial (spline) coeficient repect to the origin of coordinates
    #The range selected to get theresoult shall be 3 or 4 colums
    #and equal number of rows than the polynomials needed.
    def cerchacoef(self, lRange, keyArg, v1, v2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0
        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False
        if LenR < 2 or LenC != 2:
            cerchcoef = tuple([["Range selection is incorrect" for j in range(4)] for i in range(LenR)])
            return cerchcoef

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchcoef = tuple([["key argument larger than 3 is not allowed" for j in range(4)] for i in range(LenR)])
            return cerchcoef

        if type(v1) is str or type(v2) is str:
            cerchcoef = tuple([["v1 and v2 values must be numeric values" for j in range(4)] for i in range(LenR)])
            return cerchcoef

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchcoef = tuple([["Number of row is an even number" for j in range(4)] for i in range(LenR)])
                return cerchcoef

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchcoef = tuple([["Data must be order in ascending" for j in range(4)] for i in range(LenR)])
                return cerchcoef

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        for i in range(LenR):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        cerchcoef = tuple(Coefi)

        return cerchcoef
    #######################- END METHOD -############################


    #Function determine the area under the spline until Xs axis.
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cercharea(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cercharea = "Range selection is incorrect"
            return cercharea

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cercharea = "key argument larger than 3 is not allowed"
            return cercharea

        if type(v1) is str or type(v2) is str:
            cercharea = "v1 and v2 values must be numeric values"
            return cercharea

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cercharea = "Number of row is an even number"
                return cercharea

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cercharea = "w1 and w2 values must be numeric values"
            return cercharea

        if w2 <= w1:
            cercharea = "w2 must be higher than w1"
            return cercharea

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cercharea = "Data must be order in ascending"
                return cercharea

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        area = 0
        count = 0

        for i in range(j, k + 1):
            parea = Coefi[i][0] / 4*(w[count + 1]**4 - w[count]**4) + Coefi[i][1] / 3*(w[count + 1]**3 - w[count]**3) + Coefi[i][2] / 2*(w[count + 1]**2 - w[count]**2) + Coefi[i][3] * (w[count + 1] - w[count])
            area += parea
            count += 1

        cercharea = area

        return cercharea
    #######################- END METHOD -############################


    #Function determine the static moment under splines area respect the Xs axis.
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cerchamx(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cerchmx = "Range selection is incorrect"
            return cerchmx

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchmx = "key argument larger than 3 is not allowed"
            return cerchmx

        if type(v1) is str or type(v2) is str:
            cerchmx = "v1 and v2 values must be numeric values"
            return cerchmx

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchmx = "Number of row is an even number"
                return cerchmx

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cerchmx = "w1 and w2 values must be numeric values"
            return cerchmx

        if w2 <= w1:
            cerchmx = "w2 must be higher than w1"
            return cerchmx

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchmx = "Data must be order in ascending"
                return cerchmx

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        mx = 0
        count = 0

        for i in range(j, k + 1):
            pmx = (1 / 14 * Coefi[i][0]**2 * (w[count + 1]**7 - w[count]**7) + 1 / 6 * Coefi[i][0] * Coefi[i][1] * (w[count + 1]**6 - w[count]**6)
            + 1 / 10 * (Coefi[i][1]**2 + 2 * Coefi[i][0] * Coefi[i][2]) * (w[count + 1]**5 - w[count]**5) + 1 / 8 * (2 * Coefi[i][1] * Coefi[i][2] + 2 * Coefi[i][0] * Coefi[i][3]) * (w[count + 1]**4
            - w[count]**4) + 1 / 6 * (2 * Coefi[i][1] * Coefi[i][3] + Coefi[i][2]**2) * (w[count + 1]**3 - w[count]**3) + 1 / 2 * Coefi[i][2] * Coefi[i][3] * (w[count + 1]**2 - w[count]**2)
            + 1 / 2 * Coefi[i][3]**2 * (w[count + 1] - w[count]))
            mx += pmx
            count += 1

        cerchmx = mx

        return cerchmx
    #######################- END METHOD -############################


    #Function determine the static moment under splines area respect the Ys axis
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cerchamy(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cerchmy = "Range selection is incorrect"
            return cerchmy

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchmy = "key argument larger than 3 is not allowed"
            return cerchmy

        if type(v1) is str or type(v2) is str:
            cerchmy = "v1 and v2 values must be numeric values"
            return cerchmy

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchmy = "Number of row is an even number"
                return cerchmy

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cerchmy = "w1 and w2 values must be numeric values"
            return cerchmy

        if w2 <= w1:
            cerchmy = "w2 must be higher than w1"
            return cerchmy

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchmy = "Data must be order in ascending"
                return cerchmy

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        my = 0
        count = 0

        for i in range(j, k + 1):
            pmy = (1 / 5 * Coefi[i][0] * (w[count + 1]**5 - w[count]**5) + 1 / 4 *Coefi[i][1] * (w[count + 1]**4 - w[count]**4)
                + 1 / 3 * Coefi[i][2] * (w[count + 1]**3 - w[count]**3) + 1 / 2 * Coefi[i][3] * (w[count + 1]**2 - w[count]**2))

            my += pmy
            count += 1

        cerchmy = my

        return cerchmy
    #######################- END METHOD -############################

    #Function determine the second static inertial moment under splines area respect the Xs axis.
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cercham2x(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cerchm2x = "Range selection is incorrect"
            return cerchm2x

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchm2x = "key argument larger than 3 is not allowed"
            return cerchm2x

        if type(v1) is str or type(v2) is str:
            cerchm2x = "v1 and v2 values must be numeric values"
            return cerchm2x

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchm2x = "Number of row is an even number"
                return cerchm2x

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cerchm2x = "w1 and w2 values must be numeric values"
            return cerchm2x

        if w2 <= w1:
            cerchm2x = "w2 must be higher than w1"
            return cerchm2x

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchm2x = "Data must be order in ascending"
                return cerchm2x

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        m2x = 0
        count = 0

        for i in range(j, k + 1):
            pm2x = (-1 / 2 * Coefi[i][3]**2 * Coefi[i][2] * w[count]**2 + 1 / 2 * Coefi[i][3]**2 * Coefi[i][2] * w[count + 1]**2 - 1 / 3 * Coefi[i][2]**2 * Coefi[i][3] * w[count]**3 - 1 / 3 * Coefi[i][1] * Coefi[i][3]**2 * w[count]**3 - 1 / 4 * Coefi[i][0] * Coefi[i][3]**2 * w[count]**4 - 1 / 5 * Coefi[i][2]**2 * Coefi[i][1] * w[count]**5 - 1 / 5 * Coefi[i][3] * Coefi[i][1]**2 * w[count]**5 - 1 / 6 * Coefi[i][2] * Coefi[i][1]**2 * w[count]**6 - 1 / 6 * Coefi[i][2]**2 * Coefi[i][0] * w[count]**6
                   - 1 / 7 * Coefi[i][3] * Coefi[i][0]**2 * w[count]**7 - 1 / 8 * Coefi[i][1]**2 * Coefi[i][0] * w[count]**8 - 1 / 8 * Coefi[i][2] * Coefi[i][0]**2 * w[count]**8 - 1 / 9 * Coefi[i][1] * Coefi[i][0]**2 * w[count]**9 + 1 / 9 * Coefi[i][1] * Coefi[i][0]**2 * w[count + 1]**9 + (1 / 8 * Coefi[i][1]**2 * Coefi[i][0] + 1 / 8 * Coefi[i][2] * Coefi[i][0]**2) * w[count + 1]**8 + (1 / 21 * Coefi[i][1]**3 + 1 / 7 * Coefi[i][3] * Coefi[i][0]**2 + 2 / 7 * Coefi[i][2] * Coefi[i][1] * Coefi[i][0]) * w[count + 1]**7
                   + (1 / 6 * Coefi[i][2] * Coefi[i][1]**2 + 1 / 3 * Coefi[i][3] * Coefi[i][1] * Coefi[i][0] + 1 / 6 * Coefi[i][2]**2 * Coefi[i][0]) * w[count + 1]**6 + (1 / 5 * Coefi[i][3] * Coefi[i][1]**2 + 1 / 5 * Coefi[i][2]**2 * Coefi[i][1] + 2 / 5 * Coefi[i][0] * Coefi[i][3] * Coefi[i][2]) * w[count + 1]**5 + (1 / 4 * Coefi[i][0] * Coefi[i][3]**2 + 1 / 12 * Coefi[i][2]**3 + 1 / 2 * Coefi[i][1] * Coefi[i][3] * Coefi[i][2]) * w[count + 1]**4 + (1 / 3 * Coefi[i][2]**2 * Coefi[i][3] + 1 / 3 * Coefi[i][1] * Coefi[i][3]**2) * w[count + 1]**3 + 1 / 3 * Coefi[i][3]**3 * w[count + 1]
                   - 1 / 3 * Coefi[i][3]**3 * w[count] - 1 / 12 * Coefi[i][2]**3 * w[count]**4 - 1 / 21 * Coefi[i][1]**3 * w[count]**7 + 1 / 30 * Coefi[i][0]**3 * w[count + 1]**10 - 1 / 30 * Coefi[i][0]**3 * w[count]**10 - 1 / 2 * Coefi[i][1] * Coefi[i][3] * Coefi[i][2] * w[count]**4 - 2 / 5 * Coefi[i][0] * Coefi[i][3] * Coefi[i][2] * w[count]**5 - 1 / 3 * Coefi[i][3] * Coefi[i][1] * Coefi[i][0] * w[count]**6 - 2 / 7 * Coefi[i][2] * Coefi[i][1] * Coefi[i][0] * w[count]**7)
            m2x += pm2x
            count += 1

        cerchm2x = m2x

        return cerchm2x
    #######################- END METHOD -############################



    #Function determine the second static inertial moment under splines area respect the Ys axis.
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cercham2y(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cerchm2y = "Range selection is incorrect"
            return cerchm2y

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchm2y = "key argument larger than 3 is not allowed"
            return cerchm2y

        if type(v1) is str or type(v2) is str:
            cerchm2y = "v1 and v2 values must be numeric values"
            return cerchm2y

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchm2y = "Number of row is an even number"
                return cerchm2y

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cerchm2y = "w1 and w2 values must be numeric values"
            return cerchm2y

        if w2 <= w1:
            cerchm2y = "w2 must be higher than w1"
            return cerchm2y

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchm2y = "Data must be order in ascending"
                return cerchm2y

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        m2y = 0
        count = 0

        for i in range(j, k + 1):
            pm2y = 1 / 6 * Coefi[i][0] * (w[count + 1]**6 - w[count]**6) + 1 / 5 * Coefi[i][1] * (w[count + 1]**5 - w[count]**5) + 1 / 4 * Coefi[i][2] * (w[count + 1]**4 - w[count]**4) + 1 / 3 * Coefi[i][3] * (w[count + 1]**3 - w[count]**3)
            m2y += pm2y
            count += 1

        cerchm2y = m2y

        return cerchm2y
    #######################- END METHOD -############################

    #Function determine the inertial product under the spline with respect to the Xs and Ys axes.
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cerchap2(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cerchp2 = "Range selection is incorrect"
            return cerchp2

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchp2 = "key argument larger than 3 is not allowed"
            return cerchp2

        if type(v1) is str or type(v2) is str:
            cerchp2 = "v1 and v2 values must be numeric values"
            return cerchp2

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchp2 = "Number of row is an even number"
                return cerchp2

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cerchp2 = "w1 and w2 values must be numeric values"
            return cerchp2

        if w2 <= w1:
            cerchp2 = "w2 must be higher than w1"
            return cerchp2

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchp2 = "Data must be order in ascending"
                return cerchp2

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        p2i = 0
        count = 0

        for i in range(j, k + 1):
            pp2i = 1 / 16 * Coefi[i][0]**2 * (w[count + 1]**8  - w[count]**8) + 1 / 7 * Coefi[i][1] * Coefi[i][0] * (w[count + 1]**7 - w[count]**7) + 1 / 12 * (2 * Coefi[i][2] * Coefi[i][0] + Coefi[i][1]**2) * (w[count + 1]**6 - w[count]**6) + 1 / 10 * (2 * Coefi[i][3] * Coefi[i][0] + 2 * Coefi[i][2] * Coefi[i][1]) * (w[count + 1]**5 - w[count]**5) + 1 / 8 * (2 * Coefi[i][3] * Coefi[i][1] + Coefi[i][2]**2) * (w[count + 1]**4 - w[count]**4) + 1 / 3 * Coefi[i][3] * Coefi[i][2] * (w[count + 1]**3 - w[count]**3) + 1 / 4 * Coefi[i][3]**2 * (w[count + 1]**2 - w[count]**2)
            p2i += pp2i
            count += 1

        cerchp2 = p2i

        return cerchp2
    #######################- END METHOD -############################




    #Function determine the lengotudinal coordinate of the gravity center of the area formed under the spline.
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cerchaxg(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cerchxg = "Range selection is incorrect"
            return cerchxg

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchxg = "key argument larger than 3 is not allowed"
            return cerchxg

        if type(v1) is str or type(v2) is str:
            cerchxg = "v1 and v2 values must be numeric values"
            return cerchxg

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchxg = "Number of row is an even number"
                return cerchxg

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cerchxg = "w1 and w2 values must be numeric values"
            return cerchxg

        if w2 <= w1:
            cerchxg = "w2 must be higher than w1"
            return cerchxg

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchxg = "Data must be order in ascending"
                return cerchxg

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        my = 0
        area = 0
        count = 0

        for i in range(j, k + 1):
            pmy = (1 / 5 * Coefi[i][0] * (w[count + 1]**5 - w[count]**5) + 1 / 4 * Coefi[i][1] * (w[count + 1]**4 - w[count]**4)
                  + 1 / 3 * Coefi[i][2] * (w[count + 1]**3 - w[count]**3) + 1 / 2 * Coefi[i][3] * (w[count + 1]**2 - w[count] **2))

            parea = (Coefi[i][0] / 4 * (w[count + 1]**4 - w[count]**4) + Coefi[i][1] / 3 * (w[count + 1]**3 - w[count]**3)
                  + Coefi[i][2] / 2 * (w[count + 1]**2 - w[count]**2) + Coefi[i][3] * (w[count + 1] - w[count]))

            my += pmy
            area += parea
            count += 1

        cerchxg = my / area

        return cerchxg


    #Function determine the vertical coordinate of the gravity center of the area formed under the spline.
    #Type of spline that is needed, that is to say, settle down "end-point constraints".

    #Important: Data must be ordered in ascending and the end-point constraints will be applied, first (1st key and V1) for
    #the smaller value of lRange (1 column) and (2nd key and V2) for the greater value of lRange (1st column).
    def cerchayg(self, lRange, keyArg, v1, v2, w1, w2):
        global Coef
        global b
        global c
        global key
                             #Globla variable required for calculation
        if keyArg == None:
            keyArg = "AA"
        if v1 == None:
            v1 = 0
        if v2 == None:
            v2 = 0

        LenR = len(lRange)            #Number of rows
        LenC = len(lRange[0])          #Number of columns
        N = LenR - 1
        key = False

        if LenR < 2 or LenC != 2:
            cerchyg = "Range selection is incorrect"
            return cerchyg

        if len(keyArg) < 3:
            keyArg = keyArg.upper()
            t1 = keyArg[:1]
            t2 = keyArg[1:]
        else:
            cerchyg = "key argument larger than 3 is not allowed"
            return cerchyg

        if type(v1) is str or type(v2) is str:
            cerchyg = "v1 and v2 values must be numeric values"
            return cerchyg

        if t1 == "S" and t2 == "G":
            i = fmod(LenR, 2)
            if i == 0:
                cerchyg = "Number of row is an even number"
                return cerchyg

        if w1 == None or w1 == 0:
            w1 = lRange[0][0]

        if w2 == None or w2 == 0:
            w2 = lRange[LenR - 1][0]

        if type(w1) == str or type(w2) == str:
            cerchyg = "w1 and w2 values must be numeric values"
            return cerchyg

        if w2 <= w1:
            cerchyg = "w2 must be higher than w1"
            return cerchyg

        h = [0 for i in range(LenR)]
        for i in range(N):
            h[i] = lRange[i + 1][0] - lRange[i][0]
            if  h[i] <= 0:
                cerchyg = "Data must be order in ascending"
                return cerchyg

        Coef = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coef variable
        b = [[0] for i in range(LenR)]                                      #inicializing b variable
        c = [0 for i in range(LenR)]                                        #inicializing c variable
        did = [0 for i in range(LenR)]                                      #inicializing did variable
        Coefi = [[0 for j in range(4)] for i in range(LenR)]                 #inicializing Coefi variable

        cerchSolver(lRange, t1, t2, LenR, N, h, did, v1, v2, key, b, Coef, c)

        j = 0
        while lRange[j][0] <= w1 and j < N:
            j += 1
        if j > 0:
            j -= 1

        k = 0
        while lRange[k][0] < w2 and k <N:
            k += 1
        if k > 0:
            k -= 1

        w = [0 for i in range(LenR + 2)]

        w[0] = w1
        count = 0

        if j != k:
            for i in range(j, k):
                count += 1
                w[count] = lRange[i + 1][0]

        w[count + 1] = w2

        for i in range(LenR - 1):
            Coefi[i][0] = Coef[i][0]
            Coefi[i][1] = -3*Coef[i][0]*lRange[i][0] + Coef[i][1]
            Coefi[i][2] = 3*Coef[i][0]*lRange[i][0]**2 - 2*Coef[i][1]*lRange[i][0] + Coef[i][2]
            Coefi[i][3] = -Coef[i][0]*lRange[i][0]**3 + Coef[i][1]*lRange[i][0]**2 - Coef[i][2]*lRange[i][0] + Coef[i][3]

        mx = 0
        area = 0
        count = 0

        for i in range(j, k + 1):
            pmx = (1 / 14 * Coefi[i][0]**2 * (w[count + 1]**7 - w[count]**7) + 1 / 6 * Coefi[i][0] * Coefi[i][1] * (w[count + 1]**6 - w[count]**6)
                + 1 / 10 * (Coefi[i][1]**2 + 2 * Coefi[i][0] * Coefi[i][2]) * (w[count + 1]**5 - w[count]**5) + 1 / 8 * (2 * Coefi[i][1] * Coefi[i][2] + 2 * Coefi[i][0] * Coefi[i][3]) * (w[count + 1]**4
                - w[count]**4) + 1 / 6 * (2 * Coefi[i][1] * Coefi[i][3] + Coefi[i][2]**2) * (w[count + 1]**3 - w[count]**3) + 1 / 2 * Coefi[i][2] * Coefi[i][3] * (w[count + 1]**2 - w[count]**2)
                + 1 / 2 * Coefi[i][3]**2 * (w[count + 1] - w[count]))

            parea = (Coefi[i][0] / 4 * (w[count + 1]**4 - w[count]**4) + Coefi[i][1] / 3 * (w[count + 1]**3 - w[count]**3)
                    +Coefi[i][2] / 2 * (w[count + 1]**2 - w[count]**2) + Coefi[i][3] * (w[count + 1] - w[count]))

            mx += pmx
            area += parea
            count += 1

        cerchyg = mx / area

        return cerchyg


    #Function to call a hint help for an information
    def interpohelp(self, helpFunction):
        if helpFunction == None:
            msgbox("For complete function list and extra documentarion visit:\n"+"https://sites.google.com/view/interpolation/home")
            return "visit https://sites.google.com/view/interpolation/home"
        elif helpFunction.upper() == "INTERPO":
            return "interpo(x, XRange, YRange)"
        elif helpFunction.upper() == "INTERPO2":
            return "interpo2(x, y, Range)"
        elif helpFunction.upper() == "CERCHA":
            return "cercha(x, Range, keyArg, V1, V2)"
        else:
            msgbox("there is no "+(helpFunction[0]).upper()+" Available \n"+"For complete function list and extra documentarion visit:\n"+"https://sites.google.com/view/interpolation/home")
            return "No function with such a name"

    #######################- END METHOD -############################

###########################- END CLASS -#############################


#This function prepare the parameter to solve Cercha
def cerchSolver(lRange, t1, t2, f, N, h, did, v1, v2, keyy, bb=[[0],[0],[0]], CoefPar=[[0, 0, 0, 0],[0, 0, 0, 0]], cc=[]):
    global Coef
    global a
    global b
    global key

    if f == 2 and (t1 == "H" or t2 == "H"):         #Interpolation by cubic Hermite Spline
        Coef[0][0] = (h[0]*v2 - 2*lRange[1][1] + 2*lRange[0][1] + v1*h[0]) / h[0]**3        #mistake fixed 20181120
        Coef[0][1] = -(h[0]*v2 - 3*lRange[1][1] + 3*lRange[0][1] + 2*v1*h[0]) / h[0]**2
        Coef[0][2] = v1
        Coef[0][3] = lRange[0][1]
        c[0] = 2 * Coef[0][1]                # fixed 20181120
        c[1] = c[0] + 6*Coef[0][0]*h[0]              # fixed 20181120
        return 0

    if f == 2:
        t1 = "P"
        t2 = "G"

    if t1 == "X" or t2 == "X":
        t1 = "X"
        t2 = "X"

    if (f == 3 and t1 == "E") or (f == 3 and t2 == "E"):
        t1 = "P"            #Force it to be parabolic
        t2 = "P"

    for i in range(N):
        did[i] = (lRange[i + 1][1] - lRange[i][1]) / h[i]

    didh = [0 for i in range(f)]                    #didh inicialized
    for i in range(N - 1):
        didh[i] = did[i + 1] - did[i]

    if t1 == "P" and t2 == "G":         #First grade equation
        for i in range(N):
            c[i] = 0
            Coef[i][0] = 0
            Coef[i][1] = 0
            Coef[i][2] = did[i]
            Coef[i][3] = lRange[i][1]
        c[f - 1] = 0
        return 0

    if t1 == "S" and t2 == "G":         #Second grade equation
        for i in range(0, N - 1, 2):
            Coef[i][0] = 0
            Coef[i][1] = (-h[i + 1]*lRange[i + 1][1] + h[i + 1]*lRange[i][1] + h[i]*lRange[i + 2][1] - h[i]*lRange[i + 1][1]) / h[i + 1] / h[i] / (h[i] + h[i + 1])
            Coef[i][2] = -(-2 * h[i]*h[i + 1]*lRange[i + 1][1] + 2*h[i]*h[i + 1]*lRange[i][1] + h[i]**2 * lRange[i + 2][1] - h[i]**2 * lRange[i + 1][1] - h[i + 1]**2 * lRange[i + 1][1] + h[i + 1]**2 * lRange[i][1]) / h[i + 1] / h[i] / (h[i] + h[i + 1])
            Coef[i][3] = lRange[i][1]
            Coef[i + 1][0] = 0
            Coef[i + 1][1] = Coef[i][1]
            Coef[i + 1][2] = 2*Coef[i][1]*h[i] + Coef[i][2]
            Coef[i + 1][3] = lRange[i + 1][1]
            c[i] = 2 * Coef[i][1]
            c[i + 1] = c[i]
        c[f - 1] = c[N - 1]
        return 0

    diagb = [h[i + 1] for i in range(N - 2)]                    #Lower diagonal
    diag = [2*(h[i] + h[i + 1]) for i in range(N - 1)]          #Main diagonal
    diaga = [h[i + 1] for i in range(N - 1)]                    #Upper diagonal
    sm = [6*didh[i] for i in range(N - 1)]

    if t1 == "E":               #Extrapolated Not-a-Knot
        diag[0] =diag[0] + h[0] + h[0]**2 / h[1]
        diaga[0] = diaga[0] - h[0]**2 / h[1]
    elif t1 == "P":             #P Parabolic
        diag[0] = diag[0] + h[0]
    elif t1 == "C":             #C Curvature 2nd derivative
        sm[0] = sm[0] - h[0]*v1
    elif t1 == "F":             #F Forced 1nd derivative
        diag[0] = diag[0] - h[0] / 2
        sm[0] = sm[0] - 3*(did[0] - v1)
    elif t1 == "X":             #It is silved by Gauss-Jordan method , option X generate a matrix of five (5) diafgonals
        a = [[0 for jj in range(f - 2)] for ii in range(f - 2)]
        for i in range(N - 1):
            a[i][i], b[i][0] = diag[i], sm[i]
        for i in range(N - 2):
            a[i][i + 1], a[i + 1][i] = diaga[i], diagb[i]

        a[0][0] = a[0][0] - h[0]**2 / (2 * (h[0] + h[N - 1]))
        a[N - 2][N - 2] = a[N - 2][N - 2] - h[N - 1]**2 / (2 * (h[0] + h[N - 1]))
        a[0][N - 2] = (-1)*h[0]*h[N - 1] / (2 * (h[0] + h[N - 1]))
        a[N - 2][0] = a[0][N - 2]
        b[0][0] = b[0][0] - 3*h[0]*(did[0] - did[N - 1]) / (h[0] + h[N - 1])
        b[N - 2][0] = b[N - 2][0] - (3*h[N - 1]*(did[0] - did[N -1])/(h[0] + h[N - 1]))
        GJ(a, b)
        for i in reversed(range(2, f)):
            b[i - 1][0] = b[i - 2][0]
        b[0][0] = 3 * (did[0] - h[0]*b[1][0] / 6 - h[N - 1]*b[N - 1][0] / 6 - did[N - 1]) / (h[0] + h[N - 1])
        b[f - 1][0] = b[0][0]
        for i in range(N):
            Coef[i][0] = (b[i + 1][0] - b[i][0]) / (6 * h[i])
            Coef[i][1] = b[i][0] / 2
            Coef[i][2] = did[i] - h[i]*(2 * b[i][0] + b[i + 1][0]) / 6
            Coef[i][3] = lRange[i][1]
        key = True
        return 0

    if t2 == "E":
        diagb[N - 3] = diagb[N - 3] - h[N - 1]**2 / h[N - 2]
        diag[N - 2] = diag[N - 2] + h[N - 1] + h[N - 1]**2 / h[N - 2]
    elif t2 == "P":
        diag[N - 2] = diag[N - 2] + h[N - 1]
    elif t2 == "C":
        sm[N - 2] = sm[N - 2] - h[N - 1]*v2
    elif t2 == "F":
        diag[N - 2] = diag[N - 2] - h[N - 1] / 2
        sm[N - 2] = sm[N - 2] - 3 * (v2 - did[N - 1])
    for i in range(1, N - 1):       #Solution of the tridiagonal system
        p = diagb[i - 1] / diag[i - 1]
        diag[i] = diag[i] - p*diaga[i - 1]
        sm[i] = sm[i] - p*sm[i - 1]
    c[N - 1] = sm[N - 2] / diag[N - 2]
    for i in reversed(range(N - 2)):
        c[i + 1] = (sm[i] - diaga[i]*c[i + 2]) / diag[i]
    if t1 == "E":
        c[0] = c[1] - ((c[2] - c[1])*h[0] / h[1])
    elif t1 == "P":
        c[0] = c[1]
    elif t1 == "C":
        c[0] = v1
    elif t1 == "F":
        c[0] = 3 * (did[0] - v1) / h[0] - c[1] / 2
    else:
        c[0] = 0

    if t2 == "E":
        c[N] = c[N - 1] + ((c[N - 1] - c[N - 2]) * h[N - 1] / h[N - 2])
    elif t2 == "P":
        c[N] = c[N - 1]
    elif t2 == "C":
        c[N] = v2
    elif t2 == "F":
        c[N] = 3 * (v2 - did[N - 1]) / h[N - 1] - c[N - 1] / 2
    else:
        c[N] = 0
    for i in range(N):
        Coef[i][0] = (c[i + 1] - c[i]) / (6 * h[i])
        Coef[i][1] = c[i] / 2
        Coef[i][2] = did[i] - h[i] * (2 * c[i] + c[i + 1]) / 6
        Coef[i][3] = lRange[i][1]

    #This function make reference to Gauss Jordan "GJ" function
###########################- END FUNCTION -#############################




#Gauss-Jordan algorithm for matrix reduction with full pivot method
#A is a matrix (n x n); at the end contains the inverse of A
#B is a matrix (n x m); at the end contains the solution of AX=B
#this version apply the check for too small elements: |aij|<Tiny
#RetErr = "singular" (Det=0),  "overflow"
#rev. version of 8-12-2003
#Publicada en Internet por Leonardo Volpi  Foxes Team Piombino Italia
#http://digilander.libero.it/foxes/index.htm

def GJ(aa, bb=0, Det=0, dTiny=0, RetErr=""):

    global a
    global b

    if Det == 0:
        CalcDet = True    #detect if exsit a value for Det argument
    if bb == 0:
        mCol = 0
    else:
        mCol = len(b[0])           #Number of columns
    nRow = len(a)                  #Number of rows
    Id = [[0 for j in range(3)] for i in range(2 * nRow)]         #inicializing Id
    sw = 0                                                        #Swap counter variable inicialized
    Det = 1

    for k in range(nRow):
        #search max pivot
        iRow = k
        iCol = k
        pivotMax = 0
        for i in range(k, nRow):
            for j in range(k, nRow):
                if abs(a[i][j]) > pivotMax:
                    iRow = i
                    iCol = j
                    pivotMax = abs(a[i][j])
        #Following procedure will swap row and columns as required
        if iRow > k:
            a = SwapRow(a, k, iRow)
            if mCol > 0:
                b = SwapRow(b, k, iRow)
            if CalcDet:
                Det = -Det

            Id[sw][0] = k
            Id[sw][1] = iRow
            Id[sw][2] = 0
            sw += 1
        if iCol > k:
            a = SwapCol(a, k, iCol)
            if CalcDet:
                Det = -Det

            Id[sw][0] = k
            Id[sw][1] = iCol
            Id[sw][2] = 1
            sw += 1

        #Check pivot 0
        if abs(a[k][k]) <= dTiny:
            a[k][k] = 0
            Det = 0
            RetErr = "singular"
            return RetErr
        #normalization
        pk = a[k][k]
        if CalcDet:
            Det = Det * pk
        a[k][k] = 1
        for j in range(nRow):
            a[k][j] = a[k][j] / pk
        for j in range(mCol):
            b[k][j] = b[k][j] / pk
        #linear reduction
        for i in range(nRow):
            if i != k and a[i][k] != 0:
                pk = a[i][k]
                a[i][k] = 0
                for j in range(nRow):
                    a[i][j] = a[i][j] - pk*a[k][j]

                for j in range(mCol):
                    b[i][j] = b[i][j] - pk*b[k][j]

    #scramble rows
    for i in reversed(range(sw)):
        if Id[i][2] == 1:
            a = SwapCol(a, Id[i][0], Id[i][1])
        else:
            a = SwapRow(a, Id[i][0], Id[i][1])
            if mCol > 0:
                b = SwapRow(b, Id[i][0], Id[i][1])
###########################- END FUNCTION -#############################


#Function to swap rows BASED on published by Leonardo Volpi Foxes Team Piombino Italia
# http://digilander.libero.it/foxes/index.htm
def SwapRow(aa, k, i):
#    global a
    for j in range(len(aa[0])):
        temp = aa[i][j]
        aa[i][j] =  aa[k][j]
        aa[k][j] = temp
    return aa
###########################- END FUNCTION -#############################


#Function to swap columns BASED on published by Leonardo Volpi Foxes Team Piombino Italia
# http://digilander.libero.it/foxes/index.htm
def SwapCol(aa, k, j):
#    global a
    nRow = len(aa)
    for i in range(nRow):
        temp = aa[i][j]
        aa[i][j] = aa[i][k]
        aa[i][k] = temp

    if len(aa[0]) == (2 * nRow):
        for i in range(nRow):
            temp = aa[i][j + nRow - 1]
            aa[i][j + nRow - 1] = aa[i][k + nRow - 1]
            aa[i][k + nRow - 1] = temp
    return aa
###########################- END FUNCTION -#############################


def msgbox(message):
    ctx = uno.getComponentContext()
    sm = ctx.getServiceManager()
    toolkit = sm.createInstanceWithContext('com.sun.star.awt.Toolkit', ctx)
    MsgBox = toolkit.createMessageBox(
                                     toolkit.getDesktopWindow(),
                                     'infobox',
                                     1,
                                     'Interpolation Error',
                                     str(message))
    return MsgBox.execute()

def createInstance( ctx ):
    return InterpolationImpl( ctx )

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
    createInstance,"Astros.Montezuma.CalcTools.CalcFunctions.python.InterpolationImpl",
    ("com.sun.star.sheet.AddIn",),)
