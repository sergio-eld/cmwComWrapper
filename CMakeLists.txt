cmake_minimum_required(VERSION 3.13 FATAL_ERROR)

project(cmwComWrapper VERSION 0.1.1 LANGUAGES CXX)

# TODO: check compiler, check for COM libraries

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${PROJECT_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${PROJECT_BINARY_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${PROJECT_BINARY_DIR}/lib)

set(PUBLIC_HEADERS
	include/com_wrapper.h
	include/com_events.h
	)


add_library(${PROJECT_NAME}
	STATIC
		${PUBLIC_HEADERS}
		src/com_wrapper.cpp
		src/com_events.cpp
	)

target_include_directories(${PROJECT_NAME}
	PRIVATE
		include/${PROJECT_NAME}
	PUBLIC
		$<BUILD_INTERFACE:${PROJECT_SOURCE_DIR}/include>
		$<INSTALL_INTERFACE:include>
	)
	
set_target_properties(${PROJECT_NAME}
	PROPERTIES
		OUTPUT_NAME ${PROJECT_NAME}
		DEBUG_POSTFIX "_d"
	)
	
add_subdirectory(testing)
	
# installation

include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

set(INSTALL_CMAKEDIR ${CMAKE_INSTALL_LIBDIR}/cmake/${PROJECT_NAME})
set(targets_export_name ${PROJECT_NAME}Targets)
set(project_config ${PROJECT_BINARY_DIR}/${PROJECT_NAME}Config.cmake)
set(version_config ${PROJECT_BINARY_DIR}/${PROJECT_NAME}ConfigVergsion.cmake)

write_basic_package_version_file(
		${version_config}
	VERSION
		${PROJECT_VERSION}
	COMPATIBILITY
		SameMajorVersion
	)
	
configure_package_config_file(
		${PROJECT_SOURCE_DIR}/cmake/${PROJECT_NAME}Config.cmake.in
		${project_config}
	INSTALL_DESTINATION
		${INSTALL_CMAKEDIR}
	)

install(TARGETS
		${PROJECT_NAME}
	EXPORT
		${targets_export_name}
	ARCHIVE
		DESTINATION ${CMAKE_INSTALL_LIBDIR}
		COMPONENT lib
	LIBRARY
		DESTINATION ${CMAKE_INSTALL_LIBDIR}
		COMPONENT lib
	RUNTIME
		DESTINATION ${CMAKE_INSTALL_BINDIR}
		COMPONENT bin
	)
	
install(EXPORT
		${targets_export_name}
	DESTINATION
		${INSTALL_CMAKEDIR}
	)
	
install(FILES ${PUBLIC_HEADERS} DESTINATION ${CMAKE_INSTALL_INCLUDEDIR} COMPONENT dev)
install(FILES
		${project_config}
		${version_config}
	DESTINATION
		${INSTALL_CMAKEDIR}
	)
