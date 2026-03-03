#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>
#include <libgen.h>
#include <mach-o/dyld.h>

int main(int argc, char *argv[]) {
    // Get the path of this executable
    char exec_path[4096];
    uint32_t size = sizeof(exec_path);
    if (_NSGetExecutablePath(exec_path, &size) != 0) {
        fprintf(stderr, "ProPresenter Converter: could not get executable path\n");
        return 1;
    }

    // Resolve symlinks to get real path
    char real_path[4096];
    if (realpath(exec_path, real_path) == NULL) {
        strncpy(real_path, exec_path, sizeof(real_path) - 1);
    }

    // MacOS dir = dirname(real_path)  →  .../Contents/MacOS
    char *macos_dir = dirname(real_path);

    // Contents dir = dirname(MacOS)   →  .../Contents
    char contents_dir[4096];
    snprintf(contents_dir, sizeof(contents_dir), "%s/..", macos_dir);

    // Resources dir                   →  .../Contents/Resources
    char resources_dir[4096];
    snprintf(resources_dir, sizeof(resources_dir), "%s/Resources", contents_dir);

    // Bundled Python binary
    char python[4096];
    snprintf(python, sizeof(python),
             "%s/venv/bin/python3.13", resources_dir);

    // Boot script inside Resources
    char boot_script[4096];
    snprintf(boot_script, sizeof(boot_script),
             "%s/__boot__.py", resources_dir);

    // Keep .pyc files out of the bundle
    setenv("PYTHONDONTWRITEBYTECODE", "1", 1);

    // Execute: bundled python __boot__.py
    char *new_argv[] = { python, boot_script, NULL };
    execv(python, new_argv);

    // execv only returns on error
    fprintf(stderr, "ProPresenter Converter: failed to exec %s\n", python);
    return 1;
}
