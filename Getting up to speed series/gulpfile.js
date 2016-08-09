'use strict';

// The script requires gulp to be loaded
let gulp = require('gulp');
let server = require('gulp-express');
let ts = require('gulp-typescript');
let sass = require('gulp-sass');

// Define the "default" task > run it via: $ gulp
gulp.task('default', () => {
    console.log('Hello, from your first gulp task.');
});

// This is the "transpile" task which transpiles all TypeScript files to JavaScript
gulp.task('transpile', () => {
    return gulp.src('ts/*.ts')
               .pipe(ts())
               .pipe(gulp.dest('js'));
});

// This task transpiles SASS to CSS
gulp.task('sass', () => {
    return gulp.src('sass/*.scss')
               .pipe(sass())
               .pipe(gulp.dest('css'));
});

// This is the "watch" task, every time you change the TS file, it will call the "transpile" task
gulp.task('watch', () => {
    gulp.watch('ts/*.ts', ['transpile']);
    gulp.watch('sass/*.scss', ['sass']);
});


gulp.task('server', () => {
    // Start the server at the beginning of the task 
    server.run(['express.js']);
});