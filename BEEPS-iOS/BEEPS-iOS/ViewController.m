//
//  ViewController.m
//  BEEPS-iOS
//
//  Created by SuXinDe on 2018/3/1.
//  Copyright © 2018年 SkyPrayer Studio. All rights reserved.
//

#import "ViewController.h"
#import "ImageBEEPSDermabrasion.h"

@interface ViewController ()

@property (nonatomic, weak) IBOutlet UIImageView *originImageView;
@property (nonatomic, weak) IBOutlet UIImageView *dermabrasionImageView;

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
 
    __block UIImage *image = [UIImage imageNamed:@"mopi.jpg"];
//    __block UIImage *image = [UIImage imageNamed:@"mopi2.jpg"];
//    __block UIImage *image = [UIImage imageNamed:@"mopi3.jpg"];
//    __block UIImage *image = [UIImage imageNamed:@"mopi4.jpg"];
    self.originImageView.image = image;
    
    dispatch_async(dispatch_get_global_queue(DISPATCH_QUEUE_PRIORITY_DEFAULT, 0), ^{
        image = [ImageBEEPSDermabrasion beepsDermabrasionForImage:image];
        dispatch_async(dispatch_get_main_queue(), ^{
            self.dermabrasionImageView.image = image;
        });
    });
}


@end
