
export const formatBytes = (bytes: number, decimals = 2, format: 'Auto' | 'MB' | 'GB' | 'TB' = 'Auto') => {
    if (!bytes) return format === 'Auto' ? '0 Bytes' : `0 ${format}`;
    const k = 1024;

    if (format === 'Auto') {
        const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Number.parseFloat((bytes / Math.pow(k, i)).toFixed(decimals)) + ' ' + sizes[i];
    }

    let divisor = 1;
    switch (format) {
        case 'MB': divisor = k * k; break;
        case 'GB': divisor = k * k * k; break;
        case 'TB': divisor = k * k * k * k; break;
    }
    return (bytes / divisor).toFixed(decimals) + ' ' + format;
};
