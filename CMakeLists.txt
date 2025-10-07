# 1. 最低CMake版本
# 指定构建这个项目所需的最低CMake版本为4.0
# 如果系统上的CMake版本低于此版本，配置过程会报错
cmake_minimum_required(VERSION 4.0)

# 定义项目名称为"qingjia2"
# 这个名称会在后续的构建过程中被引用
project(qingjia2)

# 2. 基础编译配置
# 设置C++语言标准为C++20
# 编译器会按照C++20标准来编译源代码
set(CMAKE_CXX_STANDARD 20)

# 启用Qt的自动MOC（元对象编译器）处理
# MOC会处理包含Q_OBJECT宏的头文件，生成moc文件
set(CMAKE_AUTOMOC ON)

# 启用Qt的自动RCC（资源编译器）处理
# RCC会将.qrc资源文件编译成C++代码并链接到程序中
set(CMAKE_AUTORCC ON)

# 关闭Qt的自动UIC（用户界面编译器）处理
# UIC负责将.ui文件转换成C++头文件，这里选择手动处理
set(CMAKE_AUTOUIC OFF)

# 3. Qt6路径配置
# 设置Qt6的安装路径，让CMake能够找到Qt6的配置文件
# 这个路径需要根据实际的Qt安装位置进行调整
set(CMAKE_PREFIX_PATH "C:/Qt/QtMax/6.9.2/mingw_64")

# 4. 添加QXlsx依赖的Qt模块
# 查找并引入Qt6的必要组件模块
# COMPONENTS后面列出项目依赖的所有Qt模块
find_package(Qt6 COMPONENTS
        Core        # Qt核心模块，提供基本功能和非GUI类
        Gui         # 图形用户界面模块，提供窗口系统集成、事件处理等
        Widgets     # 窗口控件模块，提供按钮、文本框等UI控件
        Xml         # XML处理模块，用于读写XML文件
        REQUIRED)   # 标记这些模块为必需，如果找不到会报错并停止配置

# 5. 引入QXlsx源码目录
# 将第三方QXlsx库的源代码目录添加到构建系统中
# QXlsx是一个用于读写Excel文件的第三方Qt库
add_subdirectory(3rdparty/QXlsx-master/QXlsx)

# 6. 生成可执行文件
# 定义要构建的可执行文件目标"qingjia2"
# 并指定构建该目标所需的源文件列表
add_executable(qingjia2
        main.cpp        # 主程序入口文件，包含main函数
        mainwin.cpp     # 主窗口类的实现文件
)

# 7. 链接库
# 将必要的库链接到可执行文件目标
# 这些库提供了程序运行所需的功能
target_link_libraries(qingjia2
        Qt6::Core       # 链接Qt6核心库
        Qt6::Gui        # 链接Qt6图形界面库
        Qt6::Widgets    # 链接Qt6窗口控件库
        Qt6::Xml        # 链接Qt6 XML处理库
        QXlsx::QXlsx    # 链接第三方Excel操作库QXlsx
)

# 8. 包含头文件目录
# 将QXlsx的头文件目录添加到编译器的头文件搜索路径中
# PRIVATE表示这些头文件只用于当前目标的编译，不传递给依赖项
target_include_directories(qingjia2 PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/QXlsx-master/QXlsx/header
)

# 9. Windows下DLL和插件复制（使用Release版本）
# 只在Windows平台下执行以下代码
if (WIN32)
    # 设置Qt安装路径变量，用于后续的文件复制操作
    set(QT_INSTALL_PATH "C:/Qt/QtMax/6.9.2/mingw_64")

    # 复制Qt6Core.dll
    # 在构建完成后(POST_BUILD)执行复制命令
    # $<TARGET_FILE_DIR:${PROJECT_NAME}> 表示目标文件的输出目录
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
            COMMAND ${CMAKE_COMMAND} -E copy
            "${QT_INSTALL_PATH}/bin/Qt6Core.dll"        # 源文件路径
            "$<TARGET_FILE_DIR:${PROJECT_NAME}>"        # 目标目录
            COMMENT "Copying Qt6Core.dll"               # 执行时显示的信息
    )

    # 复制Qt6Gui.dll
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
            COMMAND ${CMAKE_COMMAND} -E copy
            "${QT_INSTALL_PATH}/bin/Qt6Gui.dll"
            "$<TARGET_FILE_DIR:${PROJECT_NAME}>"
            COMMENT "Copying Qt6Gui.dll"
    )

    # 复制Qt6Widgets.dll
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
            COMMAND ${CMAKE_COMMAND} -E copy
            "${QT_INSTALL_PATH}/bin/Qt6Widgets.dll"
            "$<TARGET_FILE_DIR:${PROJECT_NAME}>"
            COMMENT "Copying Qt6Widgets.dll"
    )

    # 复制Qt6Xml.dll
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
            COMMAND ${CMAKE_COMMAND} -E copy
            "${QT_INSTALL_PATH}/bin/Qt6Xml.dll"
            "$<TARGET_FILE_DIR:${PROJECT_NAME}>"
            COMMENT "Copying Qt6Xml.dll"
    )

    # 创建platforms目录
    # 在目标输出目录中创建platforms子目录，用于存放平台插件
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
            COMMAND ${CMAKE_COMMAND} -E make_directory
            "$<TARGET_FILE_DIR:${PROJECT_NAME}>/platforms"
            COMMENT "Creating platforms directory"
    )

    # 复制qwindows.dll平台插件
    # 这个DLL是Qt在Windows平台上的窗口系统集成插件
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
            COMMAND ${CMAKE_COMMAND} -E copy
            "${QT_INSTALL_PATH}/plugins/platforms/qwindows.dll"
            "$<TARGET_FILE_DIR:${PROJECT_NAME}>/platforms/"
            COMMENT "Copying qwindows.dll"
    )
endif()

# CPack NSIS 打包配置
# 只在Windows平台下配置安装包生成
if(WIN32)
    # 设置系统编码为UTF-8
    if(NOT DEFINED ENV{CPACK_NSIS_ENCODING})
        set(ENV{CPACK_NSIS_ENCODING} "UTF-8")
    endif()

    # 基本包信息配置
    # 设置安装包的名称
    set(CPACK_PACKAGE_NAME "患者请假条生成系统")
    # 设置软件供应商名称
    set(CPACK_PACKAGE_VENDOR "新湖镇卫生院")
    # 设置软件描述
    set(CPACK_PACKAGE_DESCRIPTION_SUMMARY "患者请假条生成系统 - 新湖镇卫生院")
    # 设置软件版本号
    set(CPACK_PACKAGE_VERSION "1.0.1")
    # 设置主版本号
    set(CPACK_PACKAGE_VERSION_MAJOR "1")
    # 设置次版本号
    set(CPACK_PACKAGE_VERSION_MINOR "0")
    # 设置修订版本号
    set(CPACK_PACKAGE_VERSION_PATCH "1")
    # 设置默认安装目录
    set(CPACK_PACKAGE_INSTALL_DIRECTORY "新湖镇卫生院\\\\患者请假条生成系统")
    # 设置可执行文件与显示名称的映射
    set(CPACK_PACKAGE_EXECUTABLES "原神" "患者请假条生成系统")

    # NSIS 特定配置
    # 指定使用NSIS作为打包工具
    set(CPACK_GENERATOR "NSIS")

    # 关键配置：启用Unicode支持
    set(CPACK_NSIS_ENCODING "UTF-8")
    set(CPACK_NSIS_UNICODE "ON")

    # 设置语言为中文
    set(CPACK_NSIS_LANGUAGE "SimpChinese")
    set(CPACK_NSIS_LANG_STRING "SimpChinese")

    # 设置安装程序显示名称
    set(CPACK_NSIS_DISPLAY_NAME "患者请假条生成系统")
    # 设置安装包名称
    set(CPACK_NSIS_PACKAGE_NAME "患者请假条生成系统")
    # 设置帮助链接URL
    set(CPACK_NSIS_HELP_LINK "https://space.bilibili.com/1854331060?spm_id_from=333.1007.0.0")
    # 设置关于信息链接URL
    set(CPACK_NSIS_URL_INFO_ABOUT "https://space.bilibili.com/1854331060?spm_id_from=333.1007.0.0")
    # 设置联系方式
    set(CPACK_NSIS_CONTACT "support@example.com")
    # 启用修改系统PATH环境变量的选项
    set(CPACK_NSIS_MODIFY_PATH ON)
    # 允许在安装新版本前卸载旧版本
    set(CPACK_NSIS_ENABLE_UNINSTALL_BEFORE_INSTALL ON)

    # 创建桌面快捷方式
    # 额外的NSIS脚本，用于创建桌面快捷方式
    set(CPACK_NSIS_CREATE_ICONS_EXTRA "
        CreateShortCut '$DESKTOP\\\\患者请假条生成系统.lnk' '$INSTDIR\\\\qingjia2.exe'
    ")
    # 卸载时删除桌面快捷方式
    set(CPACK_NSIS_DELETE_ICONS_EXTRA "
        Delete '$DESKTOP\\\\患者请假条生成系统.lnk'
    ")

    # 添加自定义NSIS脚本处理中文
    set(CPACK_NSIS_EXTRA_INSTALL_COMMANDS "
        !include 'FileFunc.nsh'
        !insertmacro GetParameters
        !insertmacro GetOptions
    ")

    # 定义需要包含的Qt运行时库文件列表
    set(QT_DLLS
            "Qt6Core.dll"       # Qt核心库
            "Qt6Gui.dll"        # Qt图形界面库
            "Qt6Widgets.dll"    # Qt窗口控件库
            "Qt6Xml.dll"        # Qt XML处理库
    )

    # 定义需要包含的MinGW运行时库文件列表
    set(MINGW_DLLS
            "libgcc_s_seh-1.dll"    # GCC异常处理运行时库
            "libstdc++-6.dll"       # C++标准库运行时
            "libwinpthread-1.dll"   # Windows线程库运行时
    )

    # 安装目标可执行文件
    # 指定将qingjia2可执行文件安装到安装目录的根目录
    install(TARGETS qingjia2
            RUNTIME DESTINATION .    # 运行时文件安装到根目录
            COMPONENT applications   # 属于applications组件
    )

    # 安装Qt DLLs
    # 遍历QT_DLLS列表中的每个DLL文件
    foreach(dll ${QT_DLLS})
        # 安装每个Qt DLL文件到安装目录的根目录
        install(FILES "${QT_INSTALL_PATH}/bin/${dll}"
                DESTINATION .        # 安装到根目录
                COMPONENT applications
        )
    endforeach()

    # 安装MinGW DLLs
    # 遍历MINGW_DLLS列表中的每个DLL文件
    foreach(dll ${MINGW_DLLS})
        # 安装每个MinGW DLL文件到安装目录的根目录
        install(FILES "${QT_INSTALL_PATH}/bin/${dll}"
                DESTINATION .        # 安装到根目录
                COMPONENT applications
        )
    endforeach()

    # 安装platforms插件目录
    # 将platforms插件目录安装到安装目录的platforms子目录
    install(DIRECTORY "${QT_INSTALL_PATH}/plugins/platforms/"
            DESTINATION "./platforms"    # 安装到platforms子目录
            COMPONENT applications
            FILES_MATCHING PATTERN "qwindows.dll"  # 只匹配qwindows.dll文件
    )

    # 复制构建后的文件到安装目录（用于本地测试）
    # 创建一个自定义目标，用于在本地测试时复制运行时文件
    add_custom_target(copy_runtime ALL
            COMMAND ${CMAKE_COMMAND} -E copy_directory
            "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/"   # 源目录：运行时输出目录
            "${CMAKE_BINARY_DIR}/install/"         # 目标目录：安装测试目录
            COMMENT "Copying runtime files to install directory"
    )

    # 设置依赖关系
    # 确保copy_runtime目标在qingjia2目标构建完成后执行
    add_dependencies(copy_runtime qingjia2)

    # 包含CPack模块
    # 这个调用会基于上面的CPACK_*变量配置来生成打包脚本
    include(CPack)
endif()