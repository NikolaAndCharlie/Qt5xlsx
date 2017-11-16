cmake_minimum_required(VERSION 3.5)
project(excel_demo)

if(CMAKE_CL_64)
   set(excel_arch x64)
else()
   set(excel_arch x86)
endif()



Set(CMAKE_CURRENT_INCLUDE ON)
Set(CMAKE_AUTOMOC ON)


find_package(Qt5Widgets REQUIRED)
find_package(Qt5Gui REQUIRED)
find_package(Qt5Core REQUIRED)
find_package(Qt5Sql REQUIRED)
find_package(Qt5Charts REQUIRED)
set(Qt_Lib Qt5::Widgets Qt5::Core Qt5::Gui Qt5::Sql Qt5::Charts)


include_directories(${CMAKE_SOURCE_DIR}/3rdparty/qt5xlsx/QtXlsx)
link_directories(${CMAKE_SOURCE_DIR}/3rdparty/qt5xlsx/${excel_arch})
set(Qt5Xlsx_lib optimized Qt5Xlsx debug Qt5Xlsxd)

file(GLOB HDR src/*.h)
file(GLOB SRC src/* .cc)

add_executable(excel_demo WIN32 ${HDR} ${SRC})
target_link_libraries(excel_demo ${Qt_Lib})
target_link_libraries(excel_demo ${Qt5Xlsx_lib})



