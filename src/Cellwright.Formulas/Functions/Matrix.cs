namespace Cellwright.Formulas.Functions;

/// <summary>
/// The small square matrix a least-squares fit needs inverted. Nothing here is general-purpose:
/// the sizes are the number of terms in a regression, which is a handful.
/// </summary>
internal static class Matrix
{
    /// <summary>
    /// Inverts a matrix by Gauss-Jordan elimination with partial pivoting.
    /// </summary>
    /// <returns>
    /// False when the matrix is singular, which for a regression means two of the explanatory
    /// columns say the same thing and no unique fit exists.
    /// </returns>
    public static bool TryInvert(double[,] matrix, out double[,] inverse)
    {
        int size = matrix.GetLength(0);
        var work = (double[,])matrix.Clone();
        inverse = new double[size, size];
        for (int i = 0; i < size; i++)
            inverse[i, i] = 1;

        // The threshold has to follow the scale of the data: cross-products of large numbers are
        // large, and a fixed epsilon would call a healthy matrix singular or a singular one fine.
        double tolerance = Magnitude(matrix) * size * 1e-14;
        for (int column = 0; column < size; column++)
        {
            int pivot = Pivot(work, column, size);
            if (Math.Abs(work[pivot, column]) <= tolerance)
                return false;

            Swap(work, column, pivot, size);
            Swap(inverse, column, pivot, size);
            Eliminate(work, inverse, column, size);
        }

        return true;
    }

    /// <summary>
    /// Determinant by Gaussian elimination with partial pivoting: the product of the pivots,
    /// with the sign flipping on every row exchange.
    /// </summary>
    public static double Determinant(double[,] matrix)
    {
        int size = matrix.GetLength(0);
        var work = (double[,])matrix.Clone();
        double result = 1;
        for (int column = 0; column < size; column++)
        {
            int pivot = Pivot(work, column, size);
            double lead = work[pivot, column];
            if (lead == 0)
                return 0;

            if (pivot != column)
            {
                Swap(work, column, pivot, size);
                result = -result;
            }

            result *= lead;
            for (int row = column + 1; row < size; row++)
            {
                double factor = work[row, column] / lead;
                if (factor == 0)
                    continue;

                for (int j = column; j < size; j++)
                    work[row, j] -= factor * work[column, j];
            }
        }

        return result;
    }

    private static int Pivot(double[,] work, int column, int size)
    {
        int pivot = column;
        for (int row = column + 1; row < size; row++)
        {
            if (Math.Abs(work[row, column]) > Math.Abs(work[pivot, column]))
                pivot = row;
        }

        return pivot;
    }

    /// <summary>Scales the pivot row to one and clears the column everywhere else.</summary>
    private static void Eliminate(double[,] work, double[,] inverse, int column, int size)
    {
        double scale = work[column, column];
        for (int j = 0; j < size; j++)
        {
            work[column, j] /= scale;
            inverse[column, j] /= scale;
        }

        for (int row = 0; row < size; row++)
        {
            double factor = work[row, column];
            if (row == column || factor == 0)
                continue;

            for (int j = 0; j < size; j++)
            {
                work[row, j] -= factor * work[column, j];
                inverse[row, j] -= factor * inverse[column, j];
            }
        }
    }

    private static void Swap(double[,] matrix, int left, int right, int size)
    {
        if (left == right)
            return;

        for (int j = 0; j < size; j++)
            (matrix[left, j], matrix[right, j]) = (matrix[right, j], matrix[left, j]);
    }

    private static double Magnitude(double[,] matrix)
    {
        double largest = 0;
        foreach (double value in matrix)
            largest = Math.Max(largest, Math.Abs(value));

        return largest == 0 ? 1 : largest;
    }
}
