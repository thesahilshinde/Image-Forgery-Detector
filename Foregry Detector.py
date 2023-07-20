#packages imported

import cv2
import xlsxwriter as wr
import numpy as np
import pandas as pd
import os
from PIL import Image
from skimage.filters import(threshold_niblack)





#function definition

list_img = []
def writexlsx(file_list, ary, n):
    path = r"Output_files_path"
    workbook = wr.Workbook(f"{path}\\threshy{n}.xlsx")
    worksheet = workbook.add_worksheet()
    title_format = workbook.add_format({'bold': True, 'align': 'center'})
    row_format = workbook.add_format({'align': 'center'})
    row = 0
    col = 0
    for m in range(n):
        worksheet.write(row, col, f"R{m}", title_format)
        col += 1
    for m in range(n):
        worksheet.write(row, col, f"G{m}", title_format)
        col += 1
    for m in range(n):
        worksheet.write(row, col, f"B{m}", title_format)
        col += 1
    worksheet.write(row, col, 'thresh_R1', title_format)
    worksheet.write(row, col + 1, 'thresh_R2', title_format)
    worksheet.write(row, col + 2, 'thresh_G1', title_format)
    worksheet.write(row, col + 3, 'thresh_G2', title_format)
    worksheet.write(row, col + 4, 'thresh_B1', title_format)
    worksheet.write(row, col + 5, 'thresh_B2', title_format)
    worksheet.write(row, col + 6, 'Label', title_format)

    row = 1
    j = 0

    # Iterate over the data and write it out row by row.
    for R1, R2, G1, G2, B1, B2 in (file_list):
        r, g, b = ary[j][0], ary[j][1], ary[j][2]
        col = 0
        for i in range(n):
            worksheet.write(row, col + i, round(r[i], 2), row_format)
        col = col + n
        for i in range(n):
            worksheet.write(row, col + i, round(g[i], 2), row_format)
        col = col + n
        for i in range(n):
            worksheet.write(row, col + i, round(b[i], 2), row_format)
        col = col + n
        worksheet.write(row, col, round(R1, 2), row_format)
        worksheet.write(row, col + 1, round(R2, 2), row_format)
        worksheet.write(row, col + 2, round(G1, 2), row_format)
        worksheet.write(row, col + 3, round(G2, 2), row_format)
        worksheet.write(row, col + 4, round(B1, 2), row_format)
        worksheet.write(row, col + 5, round(B2, 2), row_format)
        row += 1
        j += 1

    z = 1
    k = n * 3 + 6

    for i in os.listdir(r"Dataset_path"):
        if "tamp" in i:
            label = "Tampered"
        else:
            label = "Authentic"
        worksheet.write(z, k, label, row_format)
        z += 1

    workbook.close()


def thresh_niblack(li, p):
    list_img = os.listdir(p)
    arrays = []
    avgs = []
    for filename in list_img:
        im = os.path.join(p, filename)
        image = cv2.imread(im)
        imag = cv2.imread(im, cv2.IMREAD_GRAYSCALE)
        if imag is not None:
            #     for i in images:
            th3 = imag

            T = threshold_niblack(imag, window_size=25, k=0.8)

            th3[th3 < T] = 0
            th3[th3 >= T] = 255
            r, g, b = image[:, :, 2], image[:, :, 1], image[:, :, 0]  # channels R-G-B
            rt, gt, bt = th3[:, 2], th3[:, 1], th3[:, 0]  # binary image array
            r1, r2, g1, g2, b1, b2 = 0, 0, 0, 0, 0, 0  # big small
            dup = []
            cb, cw = 0, 0  # count
            for i in range(len(th3)):
                for j in range(len(th3[i])):
                    if th3[i][j] == 0:
                        r1 += r[i][j]
                        b1 += b[i][j]
                        g1 += g[i][j]
                        cb += 1
                    else:
                        r2 += r[i][j]
                        b2 += b[i][j]
                        g2 += g[i][j]
                        cw += 1
            avgs.append([r1 / cb, r2 / cw, g1 / cb, g2 / cw, b1 / cb, b2 / cw])
    return avgs


def TSBTC(li, n):
    arr = []
    for filename in os.listdir(li):
        im = os.path.join(p, filename)
        img = cv2.imread(im)
        if img is not None:
            b, g, r = img[:, :, 0], img[:, :, 1], img[:, :, 2]
            row, col = img.shape[0:2]
            b, g, r = b.flatten(), g.flatten(), r.flatten()
            b.sort(), g.sort(), r.sort()
            le = len(b)
            R, G, B = [], [], []
            for i in range(n):
                l = b[i * (le // n): (i + 1) * (le // n)]
                B.append(l.mean())

            for i in range(n):
                l = r[i * (le // n): (i + 1) * (le // n)]
                R.append(l.mean())

            for i in range(n):
                l = g[i * (le // n): (i + 1) * (le // n)]
                G.append(l.mean())

            arr.append([R, G, B])
    return arr






#Controller

count=1
p = r"Dataset_path"
li = []
for filename in os.listdir(p):
    img = os.path.join(p,filename)
    li.append(img)
    count+=1
th = thresh_niblack(li,p)

for i in range(2,11):
    ary = TSBTC(p,i)
    writexlsx(th,ary,i)
