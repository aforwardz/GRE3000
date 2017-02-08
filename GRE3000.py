#! /usr/bin/env python
# -*- coding: utf-8 -*-

import pygame
from pygame.locals import *
import xlrd
import xlutils
from xlutils.copy import copy
import time
from sys import exit

WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
RED = (255, 0, 0)
GREEN = (0, 255, 0)
BLUE = (0, 0, 255)
background_image_file1 = 'huise.jpg'
background_image_file2 = 'sencery.jpg'
excelname = u"再要你命3000.xlsx"

pygame.init()
screen = pygame.display.set_mode((640, 360), 0, 32)
pygame.display.set_caption("GRE3000 - XNY is Not You")

my_font = pygame.font.Font("/System/Library/Fonts/PingFang.ttc", 20)
word_font = pygame.font.Font("/System/Library/Fonts/PingFang.ttc", 30)
bgi1 = my_font.render(u"背景1", True, BLUE)
bgi2 = my_font.render(u"背景2", True, BLUE)
start = my_font.render(u"开始", True, BLUE)
pause = my_font.render(u"暂停", True, BLUE)
reset = my_font.render(u"重置", True, RED)

screen.fill(WHITE)
background = pygame.image.load(background_image_file1).convert()
bk = xlrd.open_workbook(excelname)
try:
    sh = bk.sheet_by_name(u"3000单词表")
except:
    print "Error"
    exit()

nrows = sh.nrows
ncols = sh.ncols
times_value = int(sh.cell_value(0, 0))
goal = word_font.render(u"是男人就定个小目标：干它50次！这是第" + unicode(times_value+1) + u"次～", True, RED)
goal_rect = goal.get_rect()
goal_rect.center = (320, 340)

row = int(sh.cell_value(1, 0))
if row is None or row == '':
    row = 0
use_time = int(sh.cell_value(2, 0))
if use_time is None or use_time == '':
    use_time = 0
time_value = my_font.render(u"已用时约" + unicode(int(use_time)) + u"小时", True, GREEN)
rank = my_font.render(u"当前级别为：不行", True, RED)
x = 0
started = 0
state = 0
start_time = []
clock = pygame.time.Clock()
cb = xlutils.copy.copy(bk)
cbs = cb.get_sheet(0)

while True:
    for event in pygame.event.get():
        if event.type == QUIT:
            end_time = time.localtime()
            if len(start_time):
                if end_time[3] < start_time[3]:
                    end_time[3] += 24
                use_time += end_time[3] - start_time[3]
            cbs.write(1, 0, row)
            cbs.write(2, 0, use_time)
            cb.save(u"再要你命3000.xlsx")
            exit()
        if event.type == KEYDOWN:
            if event.key == K_ESCAPE:
                exit()
            if event.key == K_SPACE:
                state = 1 - state
            if event.key == K_r:
                row = 0
                use_time = 0
            if event.key == K_LEFT:
                state = 0
                row -= 1
            if event.key == K_RIGHT:
                state = 0
                row += 1
        if event.type == MOUSEBUTTONDOWN:
            mouse_press = pygame.mouse.get_pressed()
            pos = pygame.mouse.get_pos()
            mouse_x = pos[0]
            mouse_y = pos[1]
            for index in range(len(mouse_press)):
                if mouse_press[index]:
                    if index == 0:
                        if mouse_x >= 10 and mouse_x <= 60 and mouse_y >= 20 and mouse_y <= 40:
                            background = pygame.image.load(background_image_file1).convert()
                        if mouse_x >= 80 and mouse_x <= 130 and mouse_y >= 20 and mouse_y <= 40:
                            background = pygame.image.load(background_image_file2).convert()
                        if mouse_x >= 450 and mouse_x <= 490 and mouse_y >= 20 and mouse_y <= 40:
                            row = 0
                            use_time = 0
                        if mouse_x >= 520 and mouse_x <= 560 and mouse_y >= 20 and mouse_y <= 40:
                            started = 1
                            if state == 0:
                                state = 1
                                start_time = time.localtime()
                        if mouse_x >= 580 and mouse_x <= 620 and mouse_y >= 20 and mouse_y <= 40:
                            state = 0

    screen.blit(background, (0, 0))
    #pygame.draw.rect(screen, GREEN, (20, 20, 70, 40))
    screen.blit(bgi1, (15, 15))
    screen.blit(bgi2, (85, 15))
    screen.blit(start, (525, 15))
    screen.blit(pause, (585, 15))
    screen.blit(reset, (455, 15))
    screen.blit(time_value, (500, 50))
    screen.blit(rank, (230, 15))
    screen.blit(goal, goal_rect)

    if started:
        if row < nrows:
            word_value = sh.cell_value(row, 1)
            mean_value = sh.cell_value(row, 2)
            if state:
                row += 1
        else:
            row = 0
            times_value += 1
            print times_value
            cbs.write(0, 0, times_value)
            cb.save(u"再要你命3000.xlsx")
            if use_time != 0:
                if use_time <= 3:
                    rank = my_font.render(u"当前级别为：猛男", True, RED)
                elif use_time <= 6:
                    rank = my_font.render(u"当前级别为：技术选手", True, RED)
                elif use_time <= 12:
                    rank = my_font.render(u"当前级别为：凑合", True, RED)
                else:
                    rank = my_font.render(u"当前级别为：不行", True, RED)
            state = 0
        row_value = word_font.render(str(row) + "/3072", True, RED)
        screen.blit(row_value, (500, 80))
        word = word_font.render(word_value, True, BLACK)
        word_rect = word.get_rect()
        word_rect.center = (300, 100)
        screen.blit(word, word_rect)
        means = mean_value.split(u'；')
        for m in range(len(means)):
            mean = word_font.render(means[m], True, BLACK)
            mean_rect = mean.get_rect()
            mean_rect.center = (300, 150 + m*30)
            screen.blit(mean, mean_rect)

        pygame.time.delay(2000)

    pygame.display.update()
    #pygame.time.delay(1000)