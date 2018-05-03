cmake_minimum_required(VERSION 3.1)

project(elevate)

file(GLOB elevate_headers src/*.h)
file(GLOB elevate_sources src/*.c src/*.rc)

add_executable(elevate ${elevate_headers} ${elevate_sources})
add_definitions(-D_UNICODE -DUNICODE)
