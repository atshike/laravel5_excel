# laravel5_excel

## 功能介绍

* laravel5 excel导出

##效果
- Daily20160918142028.xlsx

##使用

- composer.json中的require增加:"maatwebsite/excel": "~2.0.0",
- 更新：composer update
- config/app.php增加:
     'providers' => [
        Maatwebsite\Excel\ExcelServiceProvider::class,
     ]

     'aliases' => [
        'Excel' => Maatwebsite\Excel\Facades\Excel::class,
     ],
- 增加route：Route::any('/explore', 'web\AttendanceController@excel');
- < a href="{{url('/explore?month=').date("Y-m", time())}}">导出考勤< /a>


##备注
laravel5框架 Excel 导出测试
