public static class Extensions {
    public static double ToSafeDouble(this object value) {
        if ( double.TryParse(value?.ToString(), out double result) ) {
            return result;
        }
        return 0;
    }

    public static int ToSafeInt(this object value) {
        if ( int.TryParse(value?.ToString(), out int result) ) {
            return result;
        }
        return 0;
    }
}