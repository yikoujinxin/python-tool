import cv2 as cv

import numpy as np

font = cv.FONT_HERSHEY_SIMPLEX

image = cv.imread('5.png')

cv.imshow('image', image)
retval, thresholdImagre = cv.threshold(image, 100, 255, cv.THRESH_BINARY)

imgfont = cv.putText(thresholdImagre, 'thresholdImagre: 100,255', (100, 100), font, 1.2, (255, 255, 255), 2)

cv.imshow('thresholdImagre', thresholdImagre)

gray = cv.cvtColor(image, cv.COLOR_BGR2GRAY)
retval1, threshold = cv.threshold(gray, 100, 255, cv.THRESH_BINARY)
imgfont = cv.putText(threshold, 'threshold: 100,255', (100, 100), font, 1.2, (255, 255, 255), 2)
cv.imshow('threshold', threshold)

adaptiveThresholdMean = cv.adaptiveThreshold(gray,255,cv.ADAPTIVE_THRESH_MEAN_C,cv.THRESH_BINARY,13,9)

imgfont = cv.putText(adaptiveThresholdMean, 'adaptiveThresholdMean: 13,9', (100, 100), font, 1.2, (255, 255, 255), 2)

cv.imshow('adaptiveThresholdMean', adaptiveThresholdMean)

adaptiveThresholdGaua = cv.adaptiveThreshold(gray, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 13, 9)

imgfont = cv.putText(adaptiveThresholdGaua, 'adaptiveThresholdGaua: 13,9', (100, 100), font, 1.2, (255, 255, 255), 2)

cv.imshow('adaptiveThresholdGaua', adaptiveThresholdGaua)
cv.imwrite("52.png",image)