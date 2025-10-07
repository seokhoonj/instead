#include "instead.h"

/* =========================================================================
 * ReplaceColsInMat
 *
 * Replace one or more columns of a matrix with values from a vector.
 *
 * mat   : R matrix (logical, integer, double, complex, or character).
 * cols  : Integer vector of 1-based column indices to replace.
 * vec   : Vector of replacement values; length must equal
 *          nrows(mat) * length(cols), ordered by column (column-major).
 *
 * Notes:
 * - Type of vec must match mat, except integer matrices also accept logical.
 * - For numeric, logical, and complex types, values are copied element-wise.
 * - For character matrices, elements are assigned via SET_STRING_ELT.
 * - Input validity is checked (matrix type, type compatibility, length, bounds).
 * - Modifies mat in place; returns R_NilValue.
 * ========================================================================= */
SEXP ReplaceColsInMat(SEXP mat, SEXP cols, SEXP vec) {
  R_xlen_t i, j, m, ncols_replace;

  if (!isMatrix(mat))
    error("not a matrix");

  // allow logical -> integer replacement, otherwise types must match
  if (!(TYPEOF(mat) == INTSXP && TYPEOF(vec) == LGLSXP) && TYPEOF(mat) != TYPEOF(vec))
    error("different input types");

  m = nrows(mat);
  ncols_replace = XLENGTH(cols);
  R_xlen_t p = XLENGTH(vec);

  if (ncols_replace < 1)
    error("`cols` must contain at least one index");

  if (p != m * ncols_replace)
    error("length of `vec` must equal nrows(mat) * length(cols)");

  // ensure integer vector for column indices
  SEXP cols_int = PROTECT(coerceVector(cols, INTSXP));
  int *icolv = INTEGER(cols_int);

  for (j = 0; j < ncols_replace; ++j) {
    int icol = icolv[j];
    if (icol == NA_INTEGER)
      error("`cols` contains NA");
    if (icol < 1 || icol > ncols(mat))
      error("`cols` out of range");

    R_xlen_t off  = (R_xlen_t)(icol - 1) * m;  // matrix offset
    R_xlen_t voff = (R_xlen_t)j * m;           // vector offset

    switch (TYPEOF(mat)) {
    case LGLSXP: {
      int *imat = INTEGER(mat);
      const int *ivec = LOGICAL(vec);
      for (i = 0; i < m; ++i)
        imat[off + i] = ivec[voff + i];
    } break;

    case INTSXP: {
      int *imat = INTEGER(mat);
      const int *ivec = (TYPEOF(vec) == LGLSXP) ? LOGICAL(vec) : INTEGER(vec);
      for (i = 0; i < m; ++i)
        imat[off + i] = ivec[voff + i];
    } break;

    case REALSXP: {
      double *imat = REAL(mat);
      const double *ivec = REAL(vec);
      for (i = 0; i < m; ++i)
        imat[off + i] = ivec[voff + i];
    } break;

    case CPLXSXP: {
      Rcomplex *imat = COMPLEX(mat);
      const Rcomplex *ivec = COMPLEX(vec);
      for (i = 0; i < m; ++i)
        imat[off + i] = ivec[voff + i];
    } break;

    case STRSXP: {
      for (i = 0; i < m; ++i)
        SET_STRING_ELT(mat, off + i, STRING_ELT(vec, voff + i)); // NA_STRING ok
    } break;

    default:
      UNPROTECT(1);
    error("invalid input type");
    }
  }

  UNPROTECT(1);
  return R_NilValue;
}

/* =========================================================================
 * ReplaceValInMat
 *
 * Replace elements of a matrix with a constant value wherever a
 * reference matrix equals a given reference value.
 *
 * mat     : Matrix to modify (logical, integer, double, complex, or character).
 * val     : Scalar replacement value (same type as mat).
 * refmat  : Reference matrix (same type/dimensions as mat).
 * refval  : Scalar value; positions in refmat equal to this are replaced.
 *
 * Notes:
 * - Dimensions of mat and refmat must match.
 * - For numeric/logical/complex types, equality is direct.
 * - For character types, compared using strcmp().
 * - Modifies mat in place; returns R_NilValue.
 * ========================================================================= */
SEXP ReplaceValInMat(SEXP mat, SEXP val, SEXP refmat, SEXP refval) {
  // mat[refmat == refval] <- val (dim(mat) == dim(refmat))
  if (TYPEOF(mat) != TYPEOF(refmat))
    Rf_error("different input types");

  R_xlen_t i, j, m, n, col;
  m = nrows(mat);
  n = ncols(mat);

  switch (TYPEOF(mat)) {
  case LGLSXP: {
    int *imat = LOGICAL(mat);
    const int *irefmat = LOGICAL(refmat);
    int v = LOGICAL(val)[0];
    int r = LOGICAL(refval)[0];
    for (i = 0; i < n; ++i) {
      col = i * m;
      for (j = 0; j < m; ++j) {
        if (irefmat[col + j] == r)
          imat[col + j] = v;
      }
    }
  } break;

  case INTSXP: {
    int *imat = INTEGER(mat);
    const int *irefmat = INTEGER(refmat);
    int v = INTEGER(val)[0];
    int r = INTEGER(refval)[0];
    for (i = 0; i < n; ++i) {
      col = i * m;
      for (j = 0; j < m; ++j) {
        if (irefmat[col + j] == r)
          imat[col + j] = v;
      }
    }
  } break;

  case REALSXP: {
    double *imat = REAL(mat);
    const double *irefmat = REAL(refmat);
    double v = REAL(val)[0];
    double r = REAL(refval)[0];
    for (i = 0; i < n; ++i) {
      col = i * m;
      for (j = 0; j < m; ++j) {
        if (irefmat[col + j] == r)
          imat[col + j] = v;
      }
    }
  } break;

  case CPLXSXP: {
    Rcomplex *imat = COMPLEX(mat);
    const Rcomplex *irefmat = COMPLEX(refmat);
    Rcomplex v = COMPLEX(val)[0];
    Rcomplex r = COMPLEX(refval)[0];
    for (i = 0; i < n; ++i) {
      col = i * m;
      for (j = 0; j < m; ++j) {
        Rcomplex cur = irefmat[col + j];
        if (cur.r == r.r && cur.i == r.i)
          imat[col + j] = v;
      }
    }
  } break;

  case STRSXP: {
    for (i = 0; i < n; ++i) {
    col = i * m;
    for (j = 0; j < m; ++j) {
      if (!strcmp(CHAR(STRING_ELT(refmat, col + j)),
                  CHAR(STRING_ELT(refval, 0))))
        SET_STRING_ELT(mat, col + j, STRING_ELT(val, 0));
    }
  }
  } break;

  default:
    Rf_error("invalid input");
  }

  return R_NilValue;
}
