import pygame as pg
import os
from plane import *
import math

debug_text = []

def game():
    pg.init()
    display = pg.display.set_mode((500, 500))
    pg.display.set_caption('Красный барон')
    display_size = pg.display.get_surface().get_size()

    bg = pg.image.load(os.path.join('images', 'BG.jpg'))
    game_space = {'w': bg.get_width(), 'h': bg.get_height()}
    mouse_pos = [0, 0]


    player_pic = [pg.image.load(os.path.join('images', 'RedPlane.png'))]
    for i in range(len(player_pic)):
        player_pic[i] = player_pic[i].convert()
        player_pic[i].set_colorkey((255, 255, 255))

    clock = pg.time.Clock()

    def show_text():
        global debug_text
        for i in range(len(debug_text)):
            f1 = pg.font.Font(None, 20)
            text_rand = f1.render(debug_text[i], 0, (180, 0, 0))
            display.blit(text_rand, (10, i*20))
        debug_text = []

    planes = []

    def draw_window():
        display.blit(bg, (display_size[0]/2+10 - planes[0].x, display_size[1]/2-planes[0].y))

        for pl in planes:
            if planes.index(pl) == 0:
                pl.draw(display, display_size[0]/2-pl.w, display_size[1]/2-pl.h/2)

        show_text()
        pg.display.update()


    run = True
    while run:
        #создаем самолет
        if len(planes) < 1:
            planes.append(plane(display_size[0]/2+10, display_size[1]/2+10, 3, 0, player_pic[0], 'player'))

        clock.tick(30)

        for pl in planes:
            pl.move(game_space, display_size)

        mouse_angle, plane_angle, a, b = planes[0].rotation(pg, mouse_pos, display_size)

        debug_text.append('Угол мыши ' + str(round(mouse_angle*180/math.pi)))
        debug_text.append('Угол самолета ' + str(round(plane_angle*180/math.pi)))
        debug_text.append('a ' + str(round(a)))
        debug_text.append('b ' + str(round(b)))
        debug_text.append('самолет x,y,w,h ' + str(round(planes[0].x)) + '; ' + str(round(planes[0].y)) + '; '+str(planes[0].w) + '; ' +str(planes[0].h))


        draw_window()

        for event in pg.event.get():
            if event.type == pg.QUIT:
                run = False
            if event.type == pg.MOUSEMOTION:
                mouse_pos = event.pos

        debug_text.append('Позиция мыши: ' + str(mouse_pos))

if __name__ == '__main__':
    game()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
