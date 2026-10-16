<?php

/**
 * Simple PSR-4 autoloader for the Aspose\Cells_FOSS namespace. This allows the
 * library to be used without Composer.
 */
spl_autoload_register(function ($class) {
    $prefix = 'Aspose\\Cells_FOSS\\';
    if (strncmp($class, $prefix, strlen($prefix)) !== 0) {
        return;
    }

    $relative = substr($class, strlen($prefix));
    $path = __DIR__ . '/src/' . str_replace('\\', '/', $relative) . '.php';
    if (is_file($path)) {
        require $path;
    }
});
