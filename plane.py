import math as mt

class plane:
    def __init__(self, x, y, speed, angle, img, name):
        self.x = x
        self.y = y
        self.speed = speed
        self.angle = angle
        self.img = img
        self.base_img = img
        self.w = self.img.get_width()
        self.h = self.img.get_height()
        self.name = name

    def draw(self, display, x, y):
        display.blit(self.img, (x, y))

    def move(self, game_space, display_size):
        speedX = self.speed * mt.cos(self.angle)
        speedY = self.speed * mt.sin(self.angle)

        self.x = display_size[0]/2 if self.x > game_space['w'] -display_size[0]/2 else self.x
        self.x = game_space['w'] - display_size[0]/2 if self.x < display_size[0]/2 else self.x
        self.y = display_size[1]/2 if self.y > game_space['h']-display_size[1]/2 else self.y
        self.y = game_space['h'] -display_size[1]/2 if self.y < display_size[1]/2 else self.y

        self.x += speedX
        self.y += speedY

    def rotation(self, pg, mouse_pos, display_size):
        b = abs(mouse_pos[0] - 250)
        a = abs(mouse_pos[1] - 250)
        mouse_angle = mt.atan(a/b) if b != 0 else 0

        if mouse_pos[0] < display_size[0] / 2:
            if mouse_pos[1] >= display_size[1] / 2:
                mouse_angle = mt.pi - mouse_angle
            else:
                mouse_angle = mt.pi + mouse_angle
        elif mouse_pos[1] < display_size[1] / 2:
            mouse_angle = 2*mt.pi - mouse_angle

        if a == 0:
            return mouse_angle, self.angle, a, b

        #противоположный самолету угол
        alt_angle = self.angle + mt.pi
        if alt_angle > 2*mt.pi:
            alt_angle -= 2*mt.pi

        #определяем направление поворота самолета
        if self.angle < mouse_angle - 0.175 or self.angle > mouse_angle + 0.175:
            if self.angle < mt.pi:
                if mouse_angle > self.angle and mouse_angle < alt_angle:
                    self.angle += 0.0175
                else:
                    self.angle -= 0.0175
            else:
                if mouse_angle < self.angle and mouse_angle > alt_angle:
                    self.angle -= 0.0175
                else:
                    self.angle += 0.0175


        if self.angle > 2*mt.pi:
            self.angle = 0
        if self.angle < 0:
            self.angle = 2*mt.pi

        self.img = pg.transform.rotate(self.base_img, -self.angle*180/mt.pi)
        return mouse_angle, self.angle, a, b



