# -*- coding: utf-8 -*-

'''
@Time    : 2023/3/29 15:34
@Author  : fenglei
@FileName: feiji.py
@Software: PyCharm
 
'''
import pygame
import random

# 初始化 pygame 引擎
pygame.init()

# 定义游戏屏幕大小
WIDTH, HEIGHT = 480, 600

# 创建游戏窗口
screen = pygame.display.set_mode((WIDTH, HEIGHT))

# 设置游戏标题
pygame.display.set_caption("飞机大战")

# 加载游戏背景图像
bg_img = pygame.image.load("images/background.png").convert()

# 设置游戏中的音效
bullet_sound = pygame.mixer.Sound("sounds/bullet.wav")
enemy_down_sound = pygame.mixer.Sound("sounds/enemy_down.wav")
game_over_sound = pygame.mixer.Sound("sounds/game_over.wav")

# 定义游戏字体
font = pygame.font.SysFont("simhei", 30)


# 定义玩家飞机类
class Player(pygame.sprite.Sprite):
    def __init__(self):
        super().__init__()
        # 加载玩家飞机图像
        self.image = pygame.image.load("images/player.png").convert_alpha()
        self.rect = self.image.get_rect()
        self.rect.centerx = WIDTH // 2  # 初始化飞机位置
        self.rect.bottom = HEIGHT - 10
        # 定义飞机的初始速度
        self.speed = 8

        self.bullets = pygame.sprite.Group()  # 定义玩家子弹的组
        self.last_shot_time = pygame.time.get_ticks()  # 定义上次射击时间

    # 控制飞机向左移动
    def move_left(self):
        self.rect.x -= self.speed
        if self.rect.left < 0:
            self.rect.left = 0

    # 控制飞机向右移动
    def move_right(self):
        self.rect.x += self.speed
        if self.rect.right > WIDTH:
            self.rect.right = WIDTH

    # 控制飞机射击
    def shoot(self):
        now = pygame.time.get_ticks()
        if now - self.last_shot_time >= 300:
            self.last_shot_time = now
            bullet = Bullet(self.rect.centerx, self.rect.top)
            self.bullets.add(bullet)
            bullet_sound.play()


# 定义子弹类
class Bullet(pygame.sprite.Sprite):
    def __init__(self, x, y):
        super().__init__()
        self.image = pygame.image.load("images/bullet.png").convert_alpha()
        self.rect = self.image.get_rect()
        self.rect.centerx = x
        self.rect.bottom = y
        self.speed = 10

    def update(self):
        self.rect.y -= self.speed
        if self.rect.bottom < 0:
            self.kill()


# 定义敌机类
class Enemy(pygame.sprite.Sprite):
    def __init__(self):
        super().__init__()
        # 加载敌机图像
        self.image = pygame.image.load("images/enemy.png").convert_alpha()
        self.rect = self.image.get_rect()
        self.rect.x = random.randint(0, WIDTH - self.rect.width)
        self.rect.y = random.randint(-500, -50)
        # 定义敌机的速度
        self.speed = random.randint(2, 5)

        self.bullets = pygame.sprite.Group()  # 定义敌机子弹的组

    def update(self):
        # 敌机向下移动
        self.rect.y += self.speed
        if self.rect.top > HEIGHT:
            self.kill()  # 移出游戏区域，自动销毁

    # 控制敌机射击
    def shoot(self):
        bullet = Bullet(self.rect.centerx, self.rect.bottom)
        self.bullets.add(bullet)


# 定义游戏结束类
class GameOver(pygame.sprite.Sprite):
    def __init__(self, score):
        super().__init__()
        self.image = pygame.Surface((WIDTH, HEIGHT)).convert_alpha()
        self.image.fill((0, 0, 0, 0))
        self.rect = self.image.get_rect()

        # 定义游戏结束的文本信息和得分
        text_score = font.render("Your Score: {}".format(score), True, (255, 255, 255))
        text_game_over = font.render("Game Over!", True, (255, 255, 255))
        # 设置游戏结束文本的位置
        text_score_rect = text_score.get_rect()
        text_score_rect.centerx, text_score_rect.centery = WIDTH // 2, HEIGHT // 2 - 50
        text_game_over_rect = text_game_over.get_rect()
        text_game_over_rect.centerx, text_game_over_rect.centery = WIDTH // 2, HEIGHT // 2
        # 将文本信息写入到游戏结束窗口
        self.image.blit(text_score, text_score_rect)
        self.image.blit(text_game_over, text_game_over_rect)


# 创建精灵组
all_sprites = pygame.sprite.Group()  # 包含所有精灵
player_group = pygame.sprite.Group()  # 包含玩家飞机
enemy_group = pygame.sprite.Group()  # 包含敌机
bullet_group = pygame.sprite.Group()  # 包含玩家子弹和敌机子弹

# 定义玩家飞机
player = Player()
all_sprites.add(player)
player_group.add(player)

# 定义敌机
for i in range(5):
    enemy = Enemy()
    all_sprites.add(enemy)
    enemy_group.add(enemy)

# 定义游戏变量
score = 0  # 玩家得分
game_over = False  # 是否游戏结束
clock = pygame.time.Clock()

# 游戏循环
while not game_over:
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            game_over = True
        elif event.type == pygame.KEYDOWN:
            if event.key == pygame.K_LEFT:
                player.move_left()
            elif event.key == pygame.K_RIGHT:
                player.move_right()
            elif event.key == pygame.K_SPACE:
                player.shoot()

    # 更新游戏元素
    all_sprites.update()

    # 检测子弹碰撞敌机
    hits = pygame.sprite.groupcollide(bullet_group, enemy_group, True, True)
    for hit in hits:
        enemy_down_sound.play()
        score += 1
        enemy = Enemy()
        all_sprites.add(enemy)
        enemy_group.add(enemy)

    # 检测敌机碰撞玩家飞机
    enemy_hit_player = pygame.sprite.spritecollide(player, enemy_group, False)
    if enemy_hit_player:
        game_over_sound.play()
        game_over_sprite = GameOver(score)
        all_sprites.add(game_over_sprite)
        player.kill()
        game_over = True

    # 绘制游戏界面
    screen.blit(bg_img, (0, 0))
    all_sprites.draw(screen)
    # 显示得分
    text_surface = font.render("Score: {}".format(score), True, (255, 255, 255))
    screen.blit(text_surface, (10, 10))
    # 更新游戏窗口
    pygame.display.update()

    # 控制游戏帧率
    clock.tick(60)

# 退出游戏
pygame.quit()
