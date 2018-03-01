//
//  ImageBEEPS.h
//  BEEPS-iOS
//
//  Created by SuXinDe on 2018/3/1.
//  Copyright © 2018年 SkyPrayer Studio. All rights reserved.
//

/*
 References:
 http://www.cnblogs.com/Imageshop/p/3293300.html
 http://bigwww.epfl.ch/thevenaz/beeps/
 */

@import UIKit;
@import Foundation;

/**
 基于双指数边缘平滑滤波器的磨皮算法类
 */
@interface ImageBEEPSDermabrasion : NSObject

/*
 基于双指数边缘平滑滤波器的磨皮算法
 */
+ (UIImage *)beepsDermabrasionForImage:(UIImage *)image;

@end
