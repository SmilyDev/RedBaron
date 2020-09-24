class plane:
    def __init__(self, x, y, w, h, speed, angle, img):
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.speed = speed
        self.angle = angle
        self.img = pg.image.load(img)

    def show(self, display):
        # pg.drowrect(display, (0,0,0), self.x, self.y, self.w,  self.h)
        display.bilt(self.img, (self.x, self.y))

    def move():
        speedX = self.speed * mt.sin(angle)
        speedY = self.speed * mt.cos(angle)
        self.x += speedX
        self.y += speedY

    def rotation():
        self.img = pg.transform.rotate(self.img, self.angle)
