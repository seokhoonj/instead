#' Stratified sampling (by data.table groups)
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' Perform stratified sampling on a data.table, drawing a specified number
#' of rows **within each group**.
#'
#' @details
#' Let *n<sub>g</sub>* be the size of group *g*.
#'
#' * If `0 < size < 1`, the target sample size per group is computed as
#'   `method(`*n<sub>g</sub>*` * size)` where `method` is one of `"round"`, `"floor"`,
#'   or `"ceiling"`.
#' * If `size >= 1`, the target sample size per group is the same constant
#'   `size` for all groups.
#' * If `contain0 = FALSE` (default), any group whose target becomes 0 is
#'   bumped up to 1 so that **every group contributes at least one row**.
#' * Sampling is done independently within each group using
#'   `base::sample.int()` with `replace` as specified.
#'
#' The result carries an attribute `"group"` (a data.table of
#' group counts and final per-group sample sizes) for auditing.
#'
#' @param df A data.frame.
#' @param group_var Grouping columns. Supply bare names in a tidy style,
#'   e.g. `.(Species)` or `.(g1, g2)`.
#' @param size A positive number controlling per-group sample size:
#'   a proportion if `0 < size < 1`, otherwise an absolute count
#'   (same for all groups).
#' @param replace Logical; sample with replacement? Default `TRUE`.
#' @param contain0 Logical; if `TRUE`, groups with computed target 0 are
#'   allowed to contribute 0 rows. If `FALSE` (default), such groups are
#'   forced to contribute 1 row.
#' @param method Rounding method used when `0 < size < 1`.
#'   One of `"round"`, `"floor"`, or `"ceiling"`. Ignored when `size >= 1`.
#' @param verbose Logical; if `TRUE`, print a sampling summary table.
#'   Default `TRUE`.
#' @param seed Optional integer; if supplied, sets the RNG seed for
#'   reproducibility. If `NULL`, no seed is set. (Default `123`.)
#'
#' @return A data.table consisting of the sampled rows. An attribute
#'   `"group"` is attached describing per-group population size, target
#'   sample size, and realized proportion.
#'
#' @examples
#' \donttest{
#' dt <- data.table::as.data.table(iris)
#'
#' # 10% per-group sample (rounded), with replacement
#' stratified_sampling(dt, group_var = .(Species), size = 0.1, method = "round")
#'
#' # Fixed 5 rows per group (same count for all groups)
#' stratified_sampling(dt, group_var = .(Species), size = 5, replace = FALSE)
#'
#' # Allow groups to contribute zero when proportion rounds to 0
#' stratified_sampling(dt, group_var = .(Species), size = 0.01,
#'                     method = "floor", contain0 = TRUE)
#'
#' # Reproducible result with a fixed seed
#' stratified_sampling(dt, group_var = .(Species), size = 0.2, seed = 42)
#' }
#'
#' @export
stratified_sampling <- function(df, group_var, size, replace = TRUE,
                                contain0 = FALSE,
                                method = c("round", "floor", "ceiling"),
                                verbose = TRUE, seed = 123) {
  lifecycle::signal_stage("experimental", "stratified_sampling()")
  assert_class(df, "data.frame")

  env <- ensure_dt_env(df)
  dt <- env$dt

  group_var <- capture_names(dt, !!rlang::enquo(group_var))
  group <- dt[, .(n = .N), keyby = group_var]
  data.table::set(group, j = "g", value = seq_len(nrow(group)))

  if (size > 0 & size < 1) {
    method <- match.arg(method)
    data.table::set(group, j = "s", value = do.call(method, list(x = group$n * size)))
  }
  else if (size >= 1) {
    method <- "none"
    data.table::set(group, j = "s", value = size)
  }
  if (!contain0)
    data.table::set(group, i = which(group$s == 0), j = "s", value = 1)
  data.table::set(group, j = "p", value = group$s / group$n)
  if (verbose && size < 1) {
    cli::cat_rule(line = 2)
    cat(sprintf("Target prop: %.2f %% (method = %s, replace = %s)\n",
                size * 100, method, replace))
    cat(sprintf("Population : %s unit\n",
                stringr::str_pad(scales::comma(sum(group$n)), width = 14,
                                 pad = " ")))
    cat(sprintf("Sample     : %s unit\n",
                stringr::str_pad(scales::comma(sum(group$s)), width = 14,
                                 pad = " ")))
    cat(sprintf("Actual prop: %.2f %%\n", sum(group$s)/sum(group$n) * 100))
    cli::cat_rule(line = 2)
    print(cbind(group, prop = sprintf("%.2f %%", group$p * 100)))
    cli::cat_rule(line = 2)
  }
  if (nrow(group) > 1) {
    g <- NULL
    dt[group, on = group_var, `:=`(g, g)]
  }
  else {
    data.table::set(dt, j = "g", value = 1L)
  }
  n <- group$n
  s <- group$s
  spl <- split(seq_len(nrow(dt)), dt$g)
  if (!missing(seed))
    set.seed(seed)
  v <- sort(unlist(lapply(seq_along(spl), function(x) {
    if (n[x] > 1) {
      sample(spl[[x]], s[x], replace = replace)
    }
    else {
      sample(unname(spl[x]), s[x], replace = replace)
    }
  })))
  z <- dt[v]
  data.table::setorder(z, g)
  data.table::setattr(z, "group", group)
  data.table::set(z, j = "g", value = NULL)

  env$restore(z)
}

#' Randomly sample rows from a data frame
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' Draw a random sample of rows from a data frame. This is a simple wrapper
#' around [base::sample.int()] that preserves the data frame structure.
#'
#' @param x A data.frame.
#' @param size A non-negative integer giving the number of rows to sample.
#' @param replace Logical. Should sampling be with replacement? Defaults to `TRUE`.
#' @param prob A numeric vector of probability weights for sampling. Its length
#'   must equal `nrow(x)`.
#' @param seed An optional integer. If supplied, sets the random seed for
#'   reproducibility. If `NULL` (default), no seed is set.
#'
#' @return A data.frame containing the sampled rows.
#'
#' @examples
#' \donttest{
#' # Sample 5 rows with replacement
#' random_sampling(iris, size = 5)
#'
#' # Sample rows without replacement
#' random_sampling(iris, size = 5, replace = FALSE)
#'
#' # Weighted sampling
#' w <- runif(nrow(iris))
#' random_sampling(iris, size = 5, prob = w)
#'
#' # Reproducible sampling
#' random_sampling(iris, size = 5, seed = 42)
#' }
#'
#' @export
random_sampling <- function(x, size, replace = TRUE, prob = NULL, seed = NULL) {
  lifecycle::signal_stage("experimental", "random_sampling()")
  if (!inherits(x, "data.frame"))
    stop("`x` must be a data frame.", call. = FALSE)

  if (!is.null(seed)) {
    old_seed <- .Random.seed
    on.exit({
      if (exists("old_seed", inherits = FALSE)) .Random.seed <<- old_seed
    }, add = TRUE)
    set.seed(seed)
  }

  idx <- sample.int(nrow(x), size, replace, prob)
  x[idx, , drop = FALSE]
}


# Compute sample sizes ----------------------------------------------------

#' Compute required sample size for a proportion (vector or `data.table` form)
#'
#' Computes the required sample size for given true proportions (`p`),
#' relative margins of error (`r`), and confidence levels (`conf.level`)
#' using either the **Wald** (closed-form) or **Wilson** (iterative search) method.
#'
#' By default, returns a numeric vector of sample sizes.
#' If `as_dt = TRUE`, returns a tidy `data.table` with additional
#' derived statistics such as standard error and confidence bounds.
#'
#' @param p Numeric vector of true or expected proportions (0 < `p` < 1).
#' @param r Numeric vector of target relative margins of error (0 < `r` < 1).
#' @param conf.level Confidence level (default: 0.95).
#' @param method One of `c("wald", "wilson")`.
#'   - `"wald"` uses the closed-form expression
#'     \deqn{n = \frac{z^2 (1 - p)}{r^2 p}}
#'   - `"wilson"` iteratively increases `n` until the Wilson half-width ≤ \eqn{r \cdot p}.
#' @param ceiling_out Logical; if `TRUE` (default), rounds sample size up to the nearest integer.
#' @param n_upper Maximum search bound for the Wilson method (default: `1e7`).
#' @param as_dt Logical; if `TRUE`, returns a `data.table` with all parameter
#'   combinations and derived quantities.
#'
#' @return
#' - If `as_dt = FALSE`: a numeric vector of required sample sizes (same length as `p`).
#' - If `as_dt = TRUE`:  a `data.table` of class `"sample_size"` with columns:
#'   \describe{
#'     \item{p, r, conf.level}{Input parameters.}
#'     \item{n}{Required sample size.}
#'     \item{z}{Critical z-score.}
#'     \item{se}{Standard error.}
#'     \item{ci_lo, ci_hi}{Lower and upper confidence interval bounds.}
#'     \item{band_lo, band_hi}{Relative margin-of-error bands.}
#'     \item{label}{Formatted summary string.}
#'   }
#'
#' @details
#' This function provides a unified interface for sample-size estimation
#' under *relative error* constraints.
#' When `as_dt = TRUE`, the returned `"sample_size"` object can be
#' directly used with visualization helpers such as
#' `plot_sample_size()` or simulation workflows.
#'
#' @examples
#' # Vector output
#' compute_sample_size(0.1, 0.05, method = "wald")
#' compute_sample_size(0.1, 0.05, method = "wilson")
#' compute_sample_size(p = c(0.1, 0.2, 0.3), r = 0.05, method = "wald")
#'
#' # Data-table output with full parameter expansion
#' compute_sample_size(
#'   p = c(0.05, 0.1),
#'   r = c(0.1, 0.2),
#'   conf.level = c(0.95, 0.99),
#'   as_dt = TRUE
#' )
#'
#' @export
compute_sample_size <- function(p, r, conf.level = 0.95,
                                method = c("wald", "wilson"),
                                ceiling_out = TRUE,
                                n_upper = 1e7,
                                as_dt = TRUE) {
  method <- match.arg(method)

  if (!as_dt) {
    return(.compute_sample_size(
      p = p, r = r, conf.level = conf.level,
      method = method, ceiling_out = ceiling_out, n_upper = n_upper
    ))
  }

  dt <- data.table::as.data.table(expand.grid(p = p, r = r, conf.level = conf.level))
  data.table::set(
    dt, j = "n",
    value = .compute_sample_size(
      dt$p, dt$r, dt$conf.level,
      method = method, ceiling_out = ceiling_out, n_upper = n_upper
    )
  )
  data.table::set(dt, j = "z"      , value = stats::qnorm(1 - (1 - dt$conf.level) / 2))
  data.table::set(dt, j = "se"     , value = sqrt(dt$p * (1 - dt$p) / dt$n))
  data.table::set(dt, j = "ci_lo"  , value = dt$p - dt$z * dt$se)
  data.table::set(dt, j = "ci_hi"  , value = dt$p + dt$z * dt$se)
  data.table::set(dt, j = "band_lo", value = dt$p - dt$r * dt$p)
  data.table::set(dt, j = "band_hi", value = dt$p + dt$r * dt$p)
  data.table::set(
    dt, j = "label",
    value = sprintf(
      "p=%.2f%%, r=%.0f%%, cl=%.0f%%, n=%s",
      dt$p * 100, dt$r * 100, dt$conf.level * 100, instead::as_comma(dt$n)
    )
  )

  instead::prepend_class(dt, "sample_size")
}

#' Simulate sampling distributions for proportion estimates
#'
#' Given a `sample_size` object (created by `compute_sample_size(as_dt = TRUE)`),
#' this function generates simulated proportion estimates (`phat`)
#' by repeatedly sampling from the binomial distribution.
#'
#' For each unique combination of parameters (`p`, `r`, `conf.level`, `n`, etc.),
#' it draws `times` random samples of size `n` with success probability `p`
#' and computes the observed proportions (`phat = X / n`).
#'
#' The resulting table can be used to visualize or compare the sampling
#' distributions under different parameter settings.
#'
#' @param x A `sample_size` object returned by `compute_sample_size(as_dt = TRUE)`.
#' @param times Integer; number of repetitions per parameter combination
#'   (default: `5000`).
#'
#' @return A `data.table` containing:
#' \describe{
#'   \item{r, conf.level, n, z, se, ci_lo, ci_hi, band_lo, band_hi, label}{Design parameters inherited from input.}
#'   \item{phat}{Simulated sample proportions from Binomial draws.}
#' }
#'
#' @details
#' This function implements a simple Monte Carlo approximation of
#' the sampling distribution of the estimated proportion:
#' \deqn{ \hat{p} = \frac{1}{n}\mathrm{Binomial}(n, p) }
#'
#' Each parameter combination from the `sample_size` design is simulated independently.
#'
#' @seealso [compute_sample_size()]
#'
#' @examples
#' \donttest{
#' # 1. Design sample-size plan
#' ss <- compute_sample_size(p = c(0.05, 0.1), r = 0.1, as_dt = TRUE)
#'
#' # 2. Simulate sampling variability
#' simulate_sample_size(ss, times = 5000)
#' }
#'
#' @export
simulate_sample_size <- function(x, times = 5000) {
  assert_class(x, "sample_size")

  n <- p <- NULL
  .by <- c("r", "conf.level", "n", "z", "se", "ci_lo", "ci_hi",
           "band_lo", "band_hi", "label")

  x[, .(phat = stats::rbinom(times, size = n, prob = p) / n), by = .by]
}


# Intenal helper functions ------------------------------------------------

#' Compute required sample size for a proportion (relative margin)
#'
#' Unified front-end that dispatches to either the **Wald** (closed-form)
#' or **Wilson** (numerical search) method to obtain the sample size
#' achieving a target *relative margin of error* at a given confidence level.
#'
#' @param p Numeric vector of true/expected proportions (0 < `p` < 1).
#' @param r Numeric vector of target relative margins of error (0 < `r` < 1).
#' @param conf.level Confidence level (default 0.95).
#' @param method One of \code{c("wald","wilson")}. Defaults to \code{"wald"}.
#' @param ceiling_out Logical; if \code{TRUE} (default), round up to the next integer.
#'   (Used directly by Wald; passed through to Wilson.)
#' @param n_upper Maximum search bound for the Wilson method (default \code{1e7}).
#'
#' @details
#' - Wald uses the closed-form formula \deqn{n = \frac{z^2 (1 - p)}{r^2 p}}
#' - Wilson increases \eqn{n} until the Wilson half-width is no greater than
#'   \eqn{r \cdot p}. The Wald solution is used as a fast initial guess.
#'
#' @return Numeric vector of required sample sizes (same length as `p`).
#'
#' @keywords internal
.compute_sample_size <- function(p, r, conf.level = 0.95,
                                method = c("wald", "wilson"),
                                ceiling_out = TRUE,
                                n_upper = 1e7) {
  method <- match.arg(method)
  stopifnot(is.numeric(p), is.numeric(r), is.numeric(conf.level))

  # Each length must be 1 or the common L
  lp <- length(p); lr <- length(r); lc <- length(conf.level)
  max_len <- max(lp, lr, lc)

  ok_len <- function(len) len %in% c(1L, max_len)
  if (!ok_len(lp) || !ok_len(lr) || !ok_len(lc)) {
    stop(sprintf(
      "Length mismatch: lengths(p, r, conf.level) = (%d, %d, %d). Each must be 1 or the common length %d.",
      lp, lr, lc, max_len
    ), call. = FALSE)
  }

  # Recycle where needed
  if (lp == 1L) p <- rep(p, max_len)
  if (lr == 1L) r <- rep(r, max_len)
  if (lc == 1L) conf.level <- rep(conf.level, max_len)

  # (Optional) value range checks — vectorized
  if (any(!(p > 0 & p < 1))) stop("All `p` must be in (0, 1).", call. = FALSE)
  if (any(!(r > 0 & r < 1))) stop("All `r` must be in (0, 1).", call. = FALSE)
  if (any(!(conf.level > 0 & conf.level < 1))) stop("All `conf.level` must be in (0, 1).", call. = FALSE)

  if (method == "wald") {
    return(.compute_sample_size_wald(
      p = p, r = r, conf.level = conf.level, ceiling_out = ceiling_out
    ))
  }

  .compute_sample_size_wilson(
    p = p, r = r, conf.level = conf.level,
    n_upper = n_upper, ceiling_out = ceiling_out
  )
}

#' Compute required sample size (Wald method)
#'
#' Calculates the minimum sample size required to achieve a specified
#' *relative margin of error* for an estimated proportion using the
#' **Wald confidence interval** (closed-form solution).
#'
#' @param p Numeric vector of true proportions (0 < `p` < 1).
#' @param r Numeric vector of target relative margins of error (0 < `r` < 1).
#' @param conf.level Confidence level for the interval (default: 0.95).
#' @param ceiling_out Logical; if `TRUE` (default), rounds up to the next integer.
#'
#' @details
#' The sample size is computed from:
#' \deqn{n = \frac{z^2 (1 - p)}{r^2 p}}
#' where \eqn{z} is the quantile from the standard normal distribution
#' corresponding to the chosen confidence level.
#'
#' @return A numeric vector of sample sizes (same length as `p`).
#'
#' @keywords internal
.compute_sample_size_wald <- function(p, r, conf.level = 0.95, ceiling_out = TRUE) {
  stopifnot(is.numeric(p), is.numeric(r))
  # vectorized conf.level support
  z <- stats::qnorm(1 - (1 - conf.level)/2)
  n <- (z^2 * (1 - p)) / (r^2 * p)
  if (ceiling_out) ceiling(n) else n
}

#' Compute required sample size (Wilson method)
#'
#' Calculates the minimum sample size needed to achieve a specified
#' *relative margin of error* for an estimated proportion using the
#' **Wilson score confidence interval** (numerical search).
#'
#' @param p Numeric vector of true proportions (0 < `p` < 1).
#' @param r Numeric vector of target relative margins of error (0 < `r` < 1).
#' @param conf.level Confidence level for the interval (default: 0.95).
#' @param n_upper Maximum search bound for `n` (default: 1e7).
#' @param ceiling_out Logical; if `TRUE` (default), rounds up to the next integer.
#'
#' @details
#' The function iteratively increases `n` until the half-width of the Wilson
#' score interval is no greater than \code{r * p}.
#' The Wald approximation is used as an initial guess for faster convergence.
#'
#' The Wilson half-width at sample size `n` is computed as:
#' \deqn{
#'   h = \frac{z \sqrt{p(1-p)/n + z^2/(4n^2)}}{1 + z^2/n}
#' }
#'
#' @return A numeric vector of sample sizes (same length as `p`).
#' Returns `NA` if the solution exceeds `n_upper`.
#'
#' @keywords internal
.compute_sample_size_wilson <- function(p, r, conf.level = 0.95, n_upper = 1e7,
                                        ceiling_out = TRUE) {
  # vectorized conf.level support
  z <- stats::qnorm(1 - (1 - conf.level)/2)
  target <- r * p

  wilson_half_width <- function(n, p, z) {
    n <- ceiling(n)
    (z * sqrt(p*(1-p)/n + z^2/(4*n^2))) / (1 + z^2/n)
  }

  vapply(seq_along(p), function(i) {
    n0 <- max(2, .compute_sample_size_wald(p[i], r[i], conf.level[i], TRUE))
    n <- n0
    while (n <= n_upper && wilson_half_width(n, p[i], z[i]) > target[i]) {
      n <- ceiling(n * 1.05)
    }
    if (n > n_upper) NA_real_ else if (ceiling_out) ceiling(n) else n
  }, numeric(1))
}
