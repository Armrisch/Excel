<?php


namespace Excel;


use Illuminate\Support\ServiceProvider;

class ExcelServiceProvider extends ServiceProvider
{
    public function boot()
    {
        $this->publishes([
            __DIR__ . '/../config/datastructure.php' => config_path('datastructure.php')
        ]);
    }

    public function register()
    {
        $this->mergeConfigFrom(__DIR__ . '/../config/datastructure.php','datastructure.php');
        $this->app->bind(Excel::class);
    }
}