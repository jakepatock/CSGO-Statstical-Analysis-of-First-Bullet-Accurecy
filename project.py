import numpy
import math
from PIL import Image
import cv2
import statistics
from xlrd import open_workbook
from xlutils.copy import copy

xy_coordinates = []

def main():
    shot_distance_list = []

    player_name = input("Player name: ").title()
    map = input("What map was played? ")
    number_of_images = int(input("How many images? "))

    first_bullet_hit = 0
    total_duels = 0
    easy_shots = 0
    easy_shots_hit = 0
    medium_shots = 0
    medium_shots_hit = 0
    hard_shots = 0 
    hard_shots_hit = 0
    extreme_shots = 0
    extreme_shots_hit = 0

    for i in range(number_of_images):
        try:
            test_file = open(f"{i+1} y.png")
            test_file.close()
            readfile = cv2.imread(f"{i+1} y.png", 1)
            cv2.imshow('Gameimage', readfile)
            cv2.setMouseCallback('Gameimage', get_pixel_coordinate_of_head)
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            change_pixel_colors(i)
            shot_distance = distance_between_two_pixels()
            shot_distance_list.append(shot_distance)
            first_bullet_hit = first_bullet_hit + 1 
            total_duels = total_duels + 1
            easy_shots, easy_shots_hit, medium_shots, medium_shots_hit, hard_shots, hard_shots_hit, extreme_shots, extreme_shots_hit = shot_difficutly_hit_counter(shot_distance, True, easy_shots, easy_shots_hit, medium_shots, medium_shots_hit, hard_shots, hard_shots_hit, extreme_shots, extreme_shots_hit)
            
        except FileNotFoundError:
            readfile = cv2.imread(f"{i+1} n.png", 1)
            cv2.imshow('Gameimage', readfile)
            cv2.setMouseCallback('Gameimage', get_pixel_coordinate_of_head)
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            change_pixel_colors(i)
            shot_distance = distance_between_two_pixels()
            shot_distance_list.append(shot_distance)
            total_duels = total_duels + 1 
            easy_shots, easy_shots_hit, medium_shots, medium_shots_hit, hard_shots, hard_shots_hit, extreme_shots, extreme_shots_hit = shot_difficutly_hit_counter(shot_distance, False, easy_shots, easy_shots_hit, medium_shots, medium_shots_hit, hard_shots, hard_shots_hit, extreme_shots, extreme_shots_hit)
            
        except:
            print("File could not be opened")
            continue

    average_distance_from_head_var = average_distance_from_head(shot_distance_list)
    first_bullet_hit_percentage = first_shot_hit_percentage(first_bullet_hit, total_duels)
    easy_shot_hit_percentage = "Shots never taken"
    medium_shot_hit_percentage = "Shots never taken"
    hard_shots_hit_percentage = "Shots never taken"
    extreme_shots_hit_percentage = "Shots never taken"
    easy_shots_over_total_shots = "Shots never taken"
    medium_shots_over_total_shots = "Shots never taken"
    hard_shots_over_total_shots = "Shots never taken"
    extreme_shots_over_total_shots = "Shots never taken"

    try:
        easy_shot_hit_percentage = easy_shots_hit / easy_shots * 100
        medium_shot_hit_percentage = medium_shots_hit / medium_shots * 100
        hard_shots_hit_percentage = hard_shots_hit / hard_shots * 100
        extreme_shots_hit_percentage = extreme_shots_hit / extreme_shots * 100
        easy_shots_over_total_shots = easy_shots / total_duels
        medium_shots_over_total_shots = medium_shots / total_duels
        hard_shots_over_total_shots = hard_shots / total_duels
        extreme_shots_over_total_shots = extreme_shots / total_duels
    except: 
        pass

    writing_to_file(player_name, first_bullet_hit_percentage, average_distance_from_head_var, map, easy_shot_hit_percentage, medium_shot_hit_percentage, hard_shots_hit_percentage, 
    extreme_shots_hit_percentage, easy_shots_over_total_shots, medium_shots_over_total_shots, hard_shots_over_total_shots, extreme_shots_over_total_shots)


def shot_difficutly_hit_counter(shot_distance, check_hit, easy_shots, easy_shots_hit, medium_shots, medium_shots_hit, hard_shots, hard_shots_hit, extreme_shots, extreme_shots_hit):
    #easy_shot_distance = 30
    #medium_shot_distance = 100
    #hard_shot_shot_distance = 275
    #extreme over 275
    if check_hit:
        if shot_distance <= 30:
            easy_shots_hit = easy_shots_hit + 1
            easy_shots = easy_shots + 1
        if shot_distance > 30 and shot_distance <= 100 :
            medium_shots_hit = medium_shots_hit + 1
            medium_shots = medium_shots + 1
        if shot_distance > 100 and shot_distance <= 275:
            hard_shots_hit = hard_shots_hit + 1
            hard_shots = hard_shots + 1 
        if shot_distance > 275:
            extreme_shots_hit = extreme_shots_hit + 1
            extreme_shots = extreme_shots + 1 
    else:
        if shot_distance <= 30:
            easy_shots = easy_shots + 1
        if shot_distance > 30 and shot_distance <= 100 :
            medium_shots = medium_shots + 1
        if shot_distance > 100 and shot_distance <= 275:
            hard_shots = hard_shots + 1 
        if shot_distance > 275:
            extreme_shots = extreme_shots + 1 
    return easy_shots, easy_shots_hit, medium_shots, medium_shots_hit, hard_shots, hard_shots_hit, extreme_shots, extreme_shots_hit
    

def writing_to_file(player_name, first_bullet_hit_percentage, average_distance_from_head, map,easy_shot_hit_percentage, medium_shot_hit_percentage, hard_shots_hit_percentage, 
    extreme_shots_hit_percentage, easy_shots_over_total_shots, medium_shots_over_total_shots, hard_shots_over_total_shots, extreme_shots_over_total_shots): 
    #Open Workbook
    player_stats_read = open_workbook('Player Stats.xls')
    player_stats = copy(player_stats_read)
    sheet1 = player_stats.get_sheet(0)
    current_sheet = player_stats_read.sheet_by_index(0)
    row_to_write_to = current_sheet.nrows
    sheet1.write(0,0, 'Player')
    sheet1.write(0,1, 'Map')
    sheet1.write(0,2, 'First Bullet Hit Percentage (%)')  
    sheet1.write(0,3, 'Easy Shots Hit Percentage (%)')
    sheet1.write(0,4, 'Medium Shots Hit Percentage (%)') 
    sheet1.write(0,5, 'Hard Shots Hit Percentage (%)')
    sheet1.write(0,6, 'Extreme Shots Hit Percentage (%)')
    sheet1.write(0,7, 'Average Distance From Head (pixels)')
    sheet1.write(0,8, 'Easy Shots Taken Percentage (%)')
    sheet1.write(0,9, 'Medium Shots Taken Percentage (%)')
    sheet1.write(0,10, 'Hard Shots Taken Percentage (%)')
    sheet1.write(0,11, 'Extreme Shots Taken Percentage (%)')
    #wrtie varibles to workbook
    sheet1.write(row_to_write_to,0, player_name)
    sheet1.write(row_to_write_to,1, map)
    sheet1.write(row_to_write_to,2, first_bullet_hit_percentage)
    sheet1.write(row_to_write_to,3, easy_shot_hit_percentage)
    sheet1.write(row_to_write_to,4, medium_shot_hit_percentage)
    sheet1.write(row_to_write_to,5, hard_shots_hit_percentage)
    sheet1.write(row_to_write_to,6, extreme_shots_hit_percentage)
    sheet1.write(row_to_write_to,7, average_distance_from_head)
    sheet1.write(row_to_write_to,8, easy_shots_over_total_shots)
    sheet1.write(row_to_write_to,9, medium_shots_over_total_shots)
    sheet1.write(row_to_write_to,10, hard_shots_over_total_shots)
    sheet1.write(row_to_write_to,11, extreme_shots_over_total_shots)
    #save to doc
    player_stats.save('Player Stats.xls') 


def first_shot_hit_percentage(first_bullet_hit, total_duels):
    first_bullet_hit_percentage = first_bullet_hit / total_duels * 100
    return f"{first_bullet_hit_percentage}%"


def average_distance_from_head(shot_distance_list):
    average_distance_from_head_var = statistics.mean(shot_distance_list)
    return average_distance_from_head_var
    

def get_pixel_coordinate_of_head(event, x, y, flags, params):
    global xy_coordinates
    # left mouse clicks events 
    if event == cv2.EVENT_LBUTTONDOWN:
        xy_coordinates = x, y
      

def change_pixel_colors(i):
    #open image
    try:
        image = Image.open(f"{i+1} y.png").convert("RGB")
        #crosshair
        image.putpixel((960, 540), (255, 0, 170))
        #player
        image.putpixel((xy_coordinates[0], xy_coordinates[1]), (255, 212, 0))
        image.save("CSGOChanged.png")
    except FileNotFoundError:
        image = Image.open(f"{i+1} n.png").convert("RGB")
        #crosshair
        image.putpixel((960, 540), (255, 0, 170))
        #player
        image.putpixel((xy_coordinates[0], xy_coordinates[1]), (255, 212, 0))
        image.save("CSGOChanged.png")


def distance_between_two_pixels():
    #open image
    image = Image.open("CSGOChanged.png").convert("RGB")
    #convert image to array
    numpyimage = numpy.array(image)
    # #def pink pixel color
    pink = numpy.array([255, 0, 170],dtype=numpy.uint8)
    #def pink crosshair pixel location
    pinkpixel = numpy.where(numpy.all((numpyimage==pink),axis=-1))
    #def blue pixel color
    yellow = numpy.array([255, 212, 0],dtype=numpy.uint8)
    #def blue pixel location (head pixel)
    yellowpixel = numpy.where(numpy.all((numpyimage==yellow),axis=-1))
    #x axis and y axis distances the added
    distancex2 = (yellowpixel[0][0] - pinkpixel[0][0])**2
    distancey2 = (yellowpixel[1][0] - pinkpixel[1][0])**2
    distance_total = math.sqrt(distancex2 + distancey2) 
    return distance_total
    

if __name__ == "__main__":
    main()