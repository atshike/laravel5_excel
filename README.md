# laravel5_excel

## 功能介绍

* laravel5 excel导出


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


##备注
学习laravel框架、后续功能持续开发中