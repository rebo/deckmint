//! CSS-grid-inspired layout helpers for positioning slide objects.
//!
//! Rather than manually computing x/y/w/h for every object, you can define a
//! grid with column and row tracks, then call [`GridLayout::cell`] or
//! [`GridLayout::span`] to get a [`CellRect`] whose public `x`, `y`, `w`, `h`
//! fields feed directly into any builder's `.x()/.y()/.w()/.h()` setters.
//!
//! # Quick start
//!
//! ```rust,no_run
//! use deckmint::layout::{GridLayoutBuilder, GridTrack};
//!
//! // Three equal columns, 0.2" gap, 0.5" margin, 1" from top
//! let grid = GridLayoutBuilder::cols_n(3, 0.2)
//!     .origin(0.5, 1.0)
//!     .container(9.0, 4.5)
//!     .build();
//!
//! let r = grid.cell(0, 0);  // left column
//! // use r.x, r.y, r.w, r.h with any builder
//! ```

use crate::types::{Coord, PositionProps, PresLayout};

// ─── EMU per inch ────────────────────────────────────────────────────────────
const INCH: f64 = 914_400.0;

// ─── CellRect ─────────────────────────────────────────────────────────────────

/// A resolved rectangle in inches, returned by [`GridLayout::cell`] and
/// [`GridLayout::span`].
///
/// All fields are in inches. Use them directly with builder `.x()/.y()/.w()/.h()` methods.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CellRect {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

impl CellRect {
    /// Converts to a [`PositionProps`] with `Coord::Inches` values.
    pub fn to_position_props(&self) -> PositionProps {
        PositionProps {
            x: Some(Coord::Inches(self.x)),
            y: Some(Coord::Inches(self.y)),
            w: Some(Coord::Inches(self.w)),
            h: Some(Coord::Inches(self.h)),
        }
    }

    /// Shrink the rect by `amount` inches on all four sides.
    pub fn inset(&self, amount: f64) -> CellRect {
        CellRect {
            x: self.x + amount,
            y: self.y + amount,
            w: (self.w - 2.0 * amount).max(0.0),
            h: (self.h - 2.0 * amount).max(0.0),
        }
    }

    /// Shrink the rect by `horiz` inches on left/right and `vert` inches on top/bottom.
    pub fn inset_xy(&self, horiz: f64, vert: f64) -> CellRect {
        CellRect {
            x: self.x + horiz,
            y: self.y + vert,
            w: (self.w - 2.0 * horiz).max(0.0),
            h: (self.h - 2.0 * vert).max(0.0),
        }
    }

    /// Return the center point `(x, y)` in inches.
    pub fn center(&self) -> (f64, f64) {
        (self.x + self.w / 2.0, self.y + self.h / 2.0)
    }

    /// Center a fixed-size object `w × h` inside this rect. Same as [`center_in`] but as a method.
    pub fn center_object(&self, w: f64, h: f64) -> CellRect {
        center_in(self, w, h)
    }

    /// Split into left and right halves with a `gap` between them.
    pub fn halves_h(&self, gap: f64) -> (CellRect, CellRect) {
        let half_w = (self.w - gap) / 2.0;
        (
            CellRect { x: self.x, y: self.y, w: half_w, h: self.h },
            CellRect { x: self.x + half_w + gap, y: self.y, w: half_w, h: self.h },
        )
    }

    /// Split into top and bottom halves with a `gap` between them.
    pub fn halves_v(&self, gap: f64) -> (CellRect, CellRect) {
        let half_h = (self.h - gap) / 2.0;
        (
            CellRect { x: self.x, y: self.y, w: self.w, h: half_h },
            CellRect { x: self.x, y: self.y + half_h + gap, w: self.w, h: half_h },
        )
    }
}

// ─── GridTrack ────────────────────────────────────────────────────────────────

/// Sizing specification for a single grid track (column or row).
///
/// Analogous to CSS grid track sizing keywords:
/// - `Fr` → `1fr`, `2fr`, … (flexible, proportional share of remaining space)
/// - `Inches` → fixed size
/// - `Percent` → percentage of the container dimension
#[derive(Debug, Clone, Copy)]
pub enum GridTrack {
    /// Fractional unit — takes a proportional share of remaining space
    /// after fixed/percent tracks are resolved.
    Fr(f64),
    /// Fixed size in inches.
    Inches(f64),
    /// Percentage of the grid container's dimension (0.0–100.0).
    Percent(f64),
    /// At least `min` inches, then flexes up to `max_fr` share of remaining space.
    /// Analogous to CSS `minmax(min, max_fr * 1fr)`.
    MinMax { min: f64, max_fr: f64 },
}

impl GridTrack {
    pub fn fr(v: f64) -> Self { GridTrack::Fr(v) }
    pub fn inches(v: f64) -> Self { GridTrack::Inches(v) }
    pub fn pct(v: f64) -> Self { GridTrack::Percent(v) }
    /// Create `n` copies of `track` — equivalent to CSS `repeat(n, track)`.
    pub fn repeat(n: usize, track: GridTrack) -> Vec<GridTrack> { vec![track; n] }
}

// ─── Resolution algorithm ─────────────────────────────────────────────────────

/// Resolves a list of tracks into `(start_offset, size)` pairs, all in inches.
///
/// `container_size`: total space available for this axis (width or height).
/// `gap`: space between consecutive tracks in inches.
fn resolve_tracks(tracks: &[GridTrack], container_size: f64, gap: f64) -> Vec<(f64, f64)> {
    if tracks.is_empty() {
        return Vec::new();
    }
    let n = tracks.len();
    let total_gap = gap * (n.saturating_sub(1)) as f64;
    let available = container_size - total_gap;

    // First pass: resolve fixed/percent tracks and MinMax minimums; sum their sizes.
    let mut sizes = vec![0.0f64; n];
    let mut fixed_total = 0.0f64;
    let mut total_fr = 0.0f64;
    for (i, track) in tracks.iter().enumerate() {
        match track {
            GridTrack::Inches(v) => {
                sizes[i] = *v;
                fixed_total += v;
            }
            GridTrack::Percent(p) => {
                let sz = available * (p / 100.0);
                sizes[i] = sz;
                fixed_total += sz;
            }
            GridTrack::Fr(w) => {
                total_fr += w;
            }
            GridTrack::MinMax { min, max_fr } => {
                sizes[i] = *min;
                fixed_total += min;
                total_fr += max_fr;
            }
        }
    }

    // Second pass: resolve Fr and MinMax tracks from remaining space.
    let fr_space = (available - fixed_total).max(0.0);
    if total_fr > 0.0 {
        for (i, track) in tracks.iter().enumerate() {
            match track {
                GridTrack::Fr(w) => {
                    sizes[i] = (w / total_fr) * fr_space;
                }
                GridTrack::MinMax { min, max_fr } => {
                    // Add flexible portion on top of the minimum
                    sizes[i] = min + (max_fr / total_fr) * fr_space;
                }
                _ => {}
            }
        }
    }

    // Build (start_offset, size) pairs.
    let mut offsets = Vec::with_capacity(n);
    let mut cursor = 0.0f64;
    for size in &sizes {
        offsets.push((cursor, *size));
        cursor += size + gap;
    }
    offsets
}

// ─── GridLayoutBuilder ────────────────────────────────────────────────────────

/// Builder for a [`GridLayout`]. Call [`.build()`](GridLayoutBuilder::build) to
/// resolve all track sizes and produce a usable [`GridLayout`].
pub struct GridLayoutBuilder {
    cols: Vec<GridTrack>,
    rows: Vec<GridTrack>,
    col_gap: f64,
    row_gap: f64,
    container_w: f64,
    container_h: f64,
    origin_x: f64,
    origin_y: f64,
    pad_top: f64,
    pad_right: f64,
    pad_bottom: f64,
    pad_left: f64,
}

impl GridLayoutBuilder {
    /// New builder with default container = full 16:9 slide (10.0 × 5.625 inches),
    /// origin at (0, 0), no gap, single full-size cell.
    pub fn new() -> Self {
        GridLayoutBuilder {
            cols: vec![GridTrack::Fr(1.0)],
            rows: vec![GridTrack::Fr(1.0)],
            col_gap: 0.0,
            row_gap: 0.0,
            container_w: 10.0,
            container_h: 5.625,
            origin_x: 0.0,
            origin_y: 0.0,
            pad_top: 0.0,
            pad_right: 0.0,
            pad_bottom: 0.0,
            pad_left: 0.0,
        }
    }

    /// Initialise from a [`PresLayout`] so container dimensions match the slide.
    pub fn for_layout(layout: &PresLayout) -> Self {
        let mut b = GridLayoutBuilder::new();
        b.container_w = layout.width as f64 / INCH;
        b.container_h = layout.height as f64 / INCH;
        b
    }

    // ── Convenience constructors ──────────────────────────────────────────────

    /// `n` equal-width columns filling the full default slide, single row.
    pub fn cols_n(n: usize, gap: f64) -> Self {
        let n = n.max(1);
        GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0); n])
            .rows(vec![GridTrack::Fr(1.0)])
            .col_gap(gap)
    }

    /// `n` equal-height rows filling the full default slide, single column.
    pub fn rows_n(n: usize, gap: f64) -> Self {
        let n = n.max(1);
        GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Fr(1.0); n])
            .row_gap(gap)
    }

    /// `cols` × `rows` equal-size cells with uniform `gap`.
    pub fn grid_n_m(cols: usize, rows: usize, gap: f64) -> Self {
        let cols = cols.max(1);
        let rows = rows.max(1);
        GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0); cols])
            .rows(vec![GridTrack::Fr(1.0); rows])
            .gap(gap)
    }

    /// Fixed-width left sidebar + flexible right column.
    pub fn sidebar_left(sidebar_w: f64, gap: f64) -> Self {
        GridLayoutBuilder::new()
            .cols(vec![GridTrack::Inches(sidebar_w), GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Fr(1.0)])
            .col_gap(gap)
    }

    /// Fixed header + flexible content area + fixed footer rows, single column.
    pub fn header_footer(header_h: f64, footer_h: f64, gap: f64) -> Self {
        GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0)])
            .rows(vec![
                GridTrack::Inches(header_h),
                GridTrack::Fr(1.0),
                GridTrack::Inches(footer_h),
            ])
            .row_gap(gap)
    }

    // ── Builder setters ───────────────────────────────────────────────────────

    /// Override container dimensions (width × height) in inches.
    pub fn container(mut self, w: f64, h: f64) -> Self {
        self.container_w = w;
        self.container_h = h;
        self
    }

    /// Offset the grid's top-left corner from the slide origin, in inches.
    pub fn origin(mut self, x: f64, y: f64) -> Self {
        self.origin_x = x;
        self.origin_y = y;
        self
    }

    /// Set all column tracks.
    pub fn cols(mut self, tracks: Vec<GridTrack>) -> Self {
        self.cols = tracks;
        self
    }

    /// Set all row tracks.
    pub fn rows(mut self, tracks: Vec<GridTrack>) -> Self {
        self.rows = tracks;
        self
    }

    /// Set uniform gap between all columns AND rows.
    pub fn gap(mut self, inches: f64) -> Self {
        self.col_gap = inches;
        self.row_gap = inches;
        self
    }

    /// Set gap between columns only.
    pub fn col_gap(mut self, inches: f64) -> Self {
        self.col_gap = inches;
        self
    }

    /// Set gap between rows only.
    pub fn row_gap(mut self, inches: f64) -> Self {
        self.row_gap = inches;
        self
    }

    /// Uniform padding inside the container boundary on all four sides.
    pub fn padding(mut self, inches: f64) -> Self {
        self.pad_top = inches;
        self.pad_right = inches;
        self.pad_bottom = inches;
        self.pad_left = inches;
        self
    }

    /// Separate padding for each side (top, right, bottom, left).
    pub fn padding_trbl(mut self, top: f64, right: f64, bottom: f64, left: f64) -> Self {
        self.pad_top = top;
        self.pad_right = right;
        self.pad_bottom = bottom;
        self.pad_left = left;
        self
    }

    /// Resolve all track sizes and produce a [`GridLayout`].
    pub fn build(self) -> GridLayout {
        let eff_w = self.container_w - self.pad_left - self.pad_right;
        let eff_h = self.container_h - self.pad_top - self.pad_bottom;
        let col_offsets = resolve_tracks(&self.cols, eff_w, self.col_gap);
        let row_offsets = resolve_tracks(&self.rows, eff_h, self.row_gap);
        GridLayout {
            col_offsets,
            row_offsets,
            origin_x: self.origin_x + self.pad_left,
            origin_y: self.origin_y + self.pad_top,
        }
    }
}

impl Default for GridLayoutBuilder {
    fn default() -> Self { GridLayoutBuilder::new() }
}

// ─── GridLayout ───────────────────────────────────────────────────────────────

/// A resolved grid with pre-computed track positions.
///
/// Obtain via [`GridLayoutBuilder::build`].
pub struct GridLayout {
    /// `(x_start, width)` in inches for each column, relative to `origin_x`.
    col_offsets: Vec<(f64, f64)>,
    /// `(y_start, height)` in inches for each row, relative to `origin_y`.
    row_offsets: Vec<(f64, f64)>,
    origin_x: f64,
    origin_y: f64,
}

impl GridLayout {
    /// Returns the [`CellRect`] for the single cell at `(col, row)` (0-based).
    ///
    /// # Panics
    /// Panics (debug) if `col >= num_cols()` or `row >= num_rows()`.
    pub fn cell(&self, col: usize, row: usize) -> CellRect {
        self.span(col, row, 1, 1)
    }

    /// Returns the [`CellRect`] for a cell spanning `col_span` columns and
    /// `row_span` rows starting at `(col, row)`.
    ///
    /// Gap space between spanned tracks is absorbed into the returned size,
    /// matching CSS grid span behaviour.
    ///
    /// # Panics
    /// Panics (debug) if indices or spans are out of range.
    pub fn span(&self, col: usize, row: usize, col_span: usize, row_span: usize) -> CellRect {
        debug_assert!(col < self.col_offsets.len(), "col out of range");
        debug_assert!(row < self.row_offsets.len(), "row out of range");
        debug_assert!(col + col_span <= self.col_offsets.len(), "col span out of range");
        debug_assert!(row + row_span <= self.row_offsets.len(), "row span out of range");

        let (x, w) = self.col_range(col, col_span);
        let (y, h) = self.row_range(row, row_span);
        CellRect { x, y, w, h }
    }

    /// Returns `(x, w)` in absolute slide inches for columns `col..col+span`.
    pub fn col_rect(&self, col: usize, span: usize) -> (f64, f64) {
        self.col_range(col, span)
    }

    /// Returns `(y, h)` in absolute slide inches for rows `row..row+span`.
    pub fn row_rect(&self, row: usize, span: usize) -> (f64, f64) {
        self.row_range(row, span)
    }

    /// Number of defined columns.
    pub fn num_cols(&self) -> usize { self.col_offsets.len() }

    /// Number of defined rows.
    pub fn num_rows(&self) -> usize { self.row_offsets.len() }

    /// Iterate all cells in row-major order (left→right, then top→bottom).
    pub fn cells(&self) -> impl Iterator<Item = CellRect> + '_ {
        let nc = self.num_cols();
        let nr = self.num_rows();
        (0..nr).flat_map(move |r| (0..nc).map(move |c| self.cell(c, r)))
    }

    /// Iterate cells in row-major order, skipping the first `skip_rows` rows.
    /// Useful when row 0 is a header.
    pub fn content_cells(&self, skip_rows: usize) -> impl Iterator<Item = CellRect> + '_ {
        let nc = self.num_cols();
        let nr = self.num_rows();
        (skip_rows..nr).flat_map(move |r| (0..nc).map(move |c| self.cell(c, r)))
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    fn col_range(&self, col: usize, span: usize) -> (f64, f64) {
        let (start, _) = self.col_offsets[col];
        let end_col = (col + span - 1).min(self.col_offsets.len() - 1);
        let (end_start, end_w) = self.col_offsets[end_col];
        (self.origin_x + start, end_start + end_w - start)
    }

    fn row_range(&self, row: usize, span: usize) -> (f64, f64) {
        let (start, _) = self.row_offsets[row];
        let end_row = (row + span - 1).min(self.row_offsets.len() - 1);
        let (end_start, end_h) = self.row_offsets[end_row];
        (self.origin_y + start, end_start + end_h - start)
    }
}

// ─── Standalone helpers ───────────────────────────────────────────────────────

/// Centers an object of size `w × h` (inches) inside `region`.
///
/// Returns a new [`CellRect`] with all four fields set.
pub fn center_in(region: &CellRect, w: f64, h: f64) -> CellRect {
    CellRect {
        x: region.x + (region.w - w) / 2.0,
        y: region.y + (region.h - h) / 2.0,
        w,
        h,
    }
}

/// Returns the `x` coordinate for right-aligning an object of width `obj_w`
/// inside `region` with `gap` inches of padding from the right edge.
///
/// `x = region.x + region.w - obj_w - gap`
pub fn align_right(region: &CellRect, obj_w: f64, gap: f64) -> f64 {
    region.x + region.w - obj_w - gap
}

/// Returns the `x` coordinate for left-aligning with `gap` inches from the left edge.
pub fn align_left(region: &CellRect, gap: f64) -> f64 {
    region.x + gap
}

/// Returns the `y` coordinate for top-aligning with `gap` inches from the top edge.
pub fn align_top(region: &CellRect, gap: f64) -> f64 {
    region.y + gap
}

/// Returns the `y` coordinate for bottom-aligning an object of height `obj_h`
/// inside `region` with `gap` inches of padding from the bottom edge.
pub fn align_bottom(region: &CellRect, obj_h: f64, gap: f64) -> f64 {
    region.y + region.h - obj_h - gap
}

/// Splits `region` into `n` equal horizontal strips (stacked top-to-bottom)
/// with `gap` inches between them. Returns a `Vec<CellRect>` of length `n`.
pub fn split_v(region: &CellRect, n: usize, gap: f64) -> Vec<CellRect> {
    if n == 0 { return Vec::new(); }
    let total_gap = gap * (n.saturating_sub(1)) as f64;
    let strip_h = (region.h - total_gap) / n as f64;
    (0..n).map(|i| CellRect {
        x: region.x,
        y: region.y + i as f64 * (strip_h + gap),
        w: region.w,
        h: strip_h,
    }).collect()
}

/// Splits `region` into `n` equal vertical strips (arranged left-to-right)
/// with `gap` inches between them. Returns a `Vec<CellRect>` of length `n`.
pub fn split_h(region: &CellRect, n: usize, gap: f64) -> Vec<CellRect> {
    if n == 0 { return Vec::new(); }
    let total_gap = gap * (n.saturating_sub(1)) as f64;
    let strip_w = (region.w - total_gap) / n as f64;
    (0..n).map(|i| CellRect {
        x: region.x + i as f64 * (strip_w + gap),
        y: region.y,
        w: strip_w,
        h: region.h,
    }).collect()
}

// ─── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    const EPSILON: f64 = 1e-9;

    fn approx(a: f64, b: f64) -> bool { (a - b).abs() < EPSILON }

    // ── resolve_tracks ────────────────────────────────────────────────────────

    #[test]
    fn resolve_equal_fr_tracks() {
        // 3 equal Fr tracks in 9.0" with 0.0 gap → each 3.0"
        let offsets = resolve_tracks(&[GridTrack::Fr(1.0); 3], 9.0, 0.0);
        assert_eq!(offsets.len(), 3);
        assert!(approx(offsets[0].0, 0.0) && approx(offsets[0].1, 3.0));
        assert!(approx(offsets[1].0, 3.0) && approx(offsets[1].1, 3.0));
        assert!(approx(offsets[2].0, 6.0) && approx(offsets[2].1, 3.0));
    }

    #[test]
    fn resolve_fr_tracks_with_gap() {
        // 3 × Fr(1) in 9.0" with 0.3" gaps → total gap=0.6, available=9.0, track_space=8.4
        let offsets = resolve_tracks(&[GridTrack::Fr(1.0); 3], 9.0, 0.3);
        // each track = (9.0 - 0.6) / 3 = 2.8
        let expected_w = (9.0 - 0.3 * 2.0) / 3.0;
        assert!(approx(offsets[0].1, expected_w));
        assert!(approx(offsets[1].0, expected_w + 0.3));
        assert!(approx(offsets[2].0, (expected_w + 0.3) * 2.0));
    }

    #[test]
    fn resolve_mixed_fixed_and_fr() {
        // [Inches(2.0), Fr(1.0), Fr(2.0)] in 9.0" no gap
        // fixed = 2.0, fr_space = 7.0, fr(1)=7/3, fr(2)=14/3
        let tracks = vec![GridTrack::Inches(2.0), GridTrack::Fr(1.0), GridTrack::Fr(2.0)];
        let offsets = resolve_tracks(&tracks, 9.0, 0.0);
        assert!(approx(offsets[0].1, 2.0));
        assert!(approx(offsets[1].1, 7.0 / 3.0));
        assert!(approx(offsets[2].1, 14.0 / 3.0));
        assert!(approx(offsets[1].0, 2.0));
        assert!(approx(offsets[2].0, 2.0 + 7.0 / 3.0));
    }

    #[test]
    fn resolve_percent_track() {
        // [Percent(30), Fr(1)] in 10.0" no gap
        // available=10.0, percent → 3.0, fr → 7.0
        let tracks = vec![GridTrack::Percent(30.0), GridTrack::Fr(1.0)];
        let offsets = resolve_tracks(&tracks, 10.0, 0.0);
        assert!(approx(offsets[0].1, 3.0));
        assert!(approx(offsets[1].1, 7.0));
    }

    #[test]
    fn resolve_empty_tracks() {
        let offsets = resolve_tracks(&[], 9.0, 0.1);
        assert!(offsets.is_empty());
    }

    #[test]
    fn resolve_fr_space_zero_when_overflow() {
        // Fixed tracks overflow container — Fr tracks get 0, no panic.
        let tracks = vec![GridTrack::Inches(6.0), GridTrack::Inches(5.0), GridTrack::Fr(1.0)];
        let offsets = resolve_tracks(&tracks, 9.0, 0.0);
        assert!(approx(offsets[2].1, 0.0)); // no panic, just 0
    }

    // ── GridLayout::cell ──────────────────────────────────────────────────────

    #[test]
    fn cell_positions_are_correct() {
        // 3 equal columns in 9.0" with 0.2" gap, origin (0.5, 1.0)
        let grid = GridLayoutBuilder::cols_n(3, 0.2)
            .origin(0.5, 1.0)
            .container(9.0, 4.5)
            .build();
        // track_w = (9.0 - 0.4) / 3 = 2.8667"
        let tw = (9.0 - 0.2 * 2.0) / 3.0;
        let r0 = grid.cell(0, 0);
        let r1 = grid.cell(1, 0);
        let r2 = grid.cell(2, 0);
        assert!(approx(r0.x, 0.5) && approx(r0.w, tw));
        assert!(approx(r1.x, 0.5 + tw + 0.2));
        assert!(approx(r2.x, 0.5 + (tw + 0.2) * 2.0));
        assert!(approx(r0.y, 1.0) && approx(r0.h, 4.5)); // single row fills height
    }

    #[test]
    fn span_absorbs_gap() {
        // 3×1 grid, 0.2" gap. Span cols 0–1 should include the gap between them.
        let grid = GridLayoutBuilder::cols_n(3, 0.2)
            .container(9.0, 4.5)
            .build();
        let tw = (9.0 - 0.4) / 3.0;
        let s = grid.span(0, 0, 2, 1);
        // x=0, w = tw + 0.2 + tw = 2*tw + 0.2
        assert!(approx(s.x, 0.0));
        assert!(approx(s.w, 2.0 * tw + 0.2));
    }

    #[test]
    fn span_full_width() {
        let grid = GridLayoutBuilder::cols_n(3, 0.2)
            .container(9.0, 4.5)
            .build();
        let full = grid.span(0, 0, 3, 1);
        assert!(approx(full.w, 9.0));
    }

    // ── Convenience constructors ──────────────────────────────────────────────

    #[test]
    fn sidebar_left_constructor() {
        let grid = GridLayoutBuilder::sidebar_left(2.5, 0.15)
            .container(10.0, 5.625)
            .build();
        assert_eq!(grid.num_cols(), 2);
        let sidebar = grid.cell(0, 0);
        let content = grid.cell(1, 0);
        assert!(approx(sidebar.w, 2.5));
        assert!(approx(content.x, 2.5 + 0.15));
        assert!(approx(content.w, 10.0 - 2.5 - 0.15));
    }

    #[test]
    fn header_footer_constructor() {
        let grid = GridLayoutBuilder::header_footer(0.7, 0.4, 0.1)
            .container(10.0, 5.625)
            .build();
        assert_eq!(grid.num_rows(), 3);
        let header = grid.cell(0, 0);
        let footer = grid.cell(0, 2);
        let content = grid.cell(0, 1);
        assert!(approx(header.h, 0.7));
        assert!(approx(footer.h, 0.4));
        // content fills the rest: 5.625 - 0.7 - 0.4 - 0.1 - 0.1 = 4.325
        assert!(approx(content.h, 5.625 - 0.7 - 0.4 - 0.2));
    }

    #[test]
    fn grid_n_m_dimensions() {
        let grid = GridLayoutBuilder::grid_n_m(4, 3, 0.1)
            .container(8.0, 6.0)
            .build();
        assert_eq!(grid.num_cols(), 4);
        assert_eq!(grid.num_rows(), 3);
        // Each cell (col/row): (8.0 - 0.3) / 4 = 1.925 wide, (6.0 - 0.2) / 3 = 1.933 tall
        let r = grid.cell(0, 0);
        assert!(approx(r.w, (8.0 - 0.3) / 4.0));
        assert!(approx(r.h, (6.0 - 0.2) / 3.0));
    }

    // ── Standalone helpers ────────────────────────────────────────────────────

    #[test]
    fn center_in_centers_object() {
        let region = CellRect { x: 1.0, y: 2.0, w: 8.0, h: 4.0 };
        let c = center_in(&region, 2.0, 1.0);
        assert!(approx(c.x, 1.0 + 3.0)); // 1 + (8-2)/2 = 4
        assert!(approx(c.y, 2.0 + 1.5)); // 2 + (4-1)/2 = 3.5
        assert!(approx(c.w, 2.0));
        assert!(approx(c.h, 1.0));
    }

    #[test]
    fn align_right_correct() {
        let region = CellRect { x: 0.5, y: 0.0, w: 9.0, h: 1.0 };
        let x = align_right(&region, 2.0, 0.3);
        // 0.5 + 9.0 - 2.0 - 0.3 = 7.2
        assert!(approx(x, 7.2));
    }

    #[test]
    fn align_left_correct() {
        let region = CellRect { x: 0.5, y: 0.0, w: 9.0, h: 1.0 };
        assert!(approx(align_left(&region, 0.2), 0.7));
    }

    #[test]
    fn split_h_equal_strips() {
        let region = CellRect { x: 0.0, y: 0.0, w: 9.0, h: 1.0 };
        let strips = split_h(&region, 3, 0.1);
        // strip_w = (9.0 - 0.2) / 3 = 2.9333
        let sw = (9.0 - 0.2) / 3.0;
        assert_eq!(strips.len(), 3);
        assert!(approx(strips[0].x, 0.0) && approx(strips[0].w, sw));
        assert!(approx(strips[1].x, sw + 0.1));
        assert!(approx(strips[2].x, (sw + 0.1) * 2.0));
    }

    #[test]
    fn split_v_equal_strips() {
        let region = CellRect { x: 0.0, y: 1.0, w: 5.0, h: 4.0 };
        let strips = split_v(&region, 2, 0.2);
        let sh = (4.0 - 0.2) / 2.0;
        assert_eq!(strips.len(), 2);
        assert!(approx(strips[0].y, 1.0) && approx(strips[0].h, sh));
        assert!(approx(strips[1].y, 1.0 + sh + 0.2));
    }

    #[test]
    fn split_returns_empty_for_zero() {
        let region = CellRect { x: 0.0, y: 0.0, w: 9.0, h: 5.0 };
        assert!(split_h(&region, 0, 0.1).is_empty());
        assert!(split_v(&region, 0, 0.1).is_empty());
    }

    // ── New: CellRect convenience methods ──────────────────────────────────

    #[test]
    fn cellrect_inset() {
        let r = CellRect { x: 1.0, y: 2.0, w: 6.0, h: 4.0 };
        let i = r.inset(0.5);
        assert!(approx(i.x, 1.5) && approx(i.y, 2.5));
        assert!(approx(i.w, 5.0) && approx(i.h, 3.0));
    }

    #[test]
    fn cellrect_inset_clamps_to_zero() {
        let r = CellRect { x: 0.0, y: 0.0, w: 1.0, h: 1.0 };
        let i = r.inset(2.0);
        assert!(approx(i.w, 0.0) && approx(i.h, 0.0));
    }

    #[test]
    fn cellrect_center() {
        let r = CellRect { x: 1.0, y: 2.0, w: 8.0, h: 4.0 };
        let (cx, cy) = r.center();
        assert!(approx(cx, 5.0) && approx(cy, 4.0));
    }

    #[test]
    fn cellrect_halves_h() {
        let r = CellRect { x: 0.0, y: 0.0, w: 9.0, h: 4.0 };
        let (left, right) = r.halves_h(0.2);
        assert!(approx(left.w, 4.4) && approx(right.w, 4.4));
        assert!(approx(right.x, 4.6));
    }

    #[test]
    fn cellrect_halves_v() {
        let r = CellRect { x: 0.0, y: 1.0, w: 9.0, h: 4.0 };
        let (top, bot) = r.halves_v(0.2);
        assert!(approx(top.h, 1.9) && approx(bot.h, 1.9));
        assert!(approx(bot.y, 3.1));
    }

    // ── New: align_top / align_bottom ────────────────────────────────────

    #[test]
    fn align_top_correct() {
        let region = CellRect { x: 0.0, y: 1.0, w: 9.0, h: 4.0 };
        assert!(approx(align_top(&region, 0.2), 1.2));
    }

    #[test]
    fn align_bottom_correct() {
        let region = CellRect { x: 0.0, y: 1.0, w: 9.0, h: 4.0 };
        assert!(approx(align_bottom(&region, 0.5, 0.2), 4.3)); // 1+4 - 0.5 - 0.2
    }

    // ── New: padding ─────────────────────────────────────────────────────

    #[test]
    fn padding_shrinks_container() {
        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Fr(1.0)])
            .container(10.0, 6.0)
            .origin(0.0, 0.0)
            .padding(0.5)
            .build();
        let r = grid.cell(0, 0);
        assert!(approx(r.x, 0.5) && approx(r.y, 0.5));
        assert!(approx(r.w, 9.0) && approx(r.h, 5.0));
    }

    #[test]
    fn padding_trbl() {
        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Fr(1.0)])
            .container(10.0, 6.0)
            .padding_trbl(1.0, 0.5, 0.5, 1.0)
            .build();
        let r = grid.cell(0, 0);
        assert!(approx(r.x, 1.0) && approx(r.y, 1.0));
        assert!(approx(r.w, 8.5) && approx(r.h, 4.5));
    }

    // ── New: MinMax track ────────────────────────────────────────────────

    #[test]
    fn minmax_track_basic() {
        // [MinMax{1.0, 1.0}, Fr(1.0)] in 10.0" no gap
        // First pass: min=1.0 fixed, total_fr=2.0
        // Second pass: fr_space = 10-1 = 9.0, MinMax gets 1.0 + 0.5*9=5.5, Fr gets 0.5*9=4.5
        let tracks = vec![
            GridTrack::MinMax { min: 1.0, max_fr: 1.0 },
            GridTrack::Fr(1.0),
        ];
        let offsets = resolve_tracks(&tracks, 10.0, 0.0);
        assert!(approx(offsets[0].1, 5.5));
        assert!(approx(offsets[1].1, 4.5));
    }

    #[test]
    fn minmax_gets_at_least_min_on_overflow() {
        // When fixed tracks eat all space, MinMax still gets its min
        let tracks = vec![
            GridTrack::Inches(9.0),
            GridTrack::MinMax { min: 2.0, max_fr: 1.0 },
        ];
        let offsets = resolve_tracks(&tracks, 10.0, 0.0);
        // fixed_total = 9+2 = 11, fr_space=max(10-11,0)=0, MinMax=2.0+0=2.0
        assert!(approx(offsets[1].1, 2.0));
    }

    // ── New: repeat ──────────────────────────────────────────────────────

    #[test]
    fn repeat_creates_n_copies() {
        let tracks = GridTrack::repeat(4, GridTrack::Fr(1.0));
        assert_eq!(tracks.len(), 4);
    }

    // ── New: cells() iterator ────────────────────────────────────────────

    #[test]
    fn cells_iterator_row_major() {
        let grid = GridLayoutBuilder::grid_n_m(3, 2, 0.0)
            .container(9.0, 4.0)
            .build();
        let cells: Vec<CellRect> = grid.cells().collect();
        assert_eq!(cells.len(), 6);
        // First cell is (0,0), second is (1,0), third is (2,0), fourth is (0,1)...
        assert!(approx(cells[0].x, 0.0) && approx(cells[0].y, 0.0));
        assert!(approx(cells[3].x, 0.0) && approx(cells[3].y, 2.0));
    }

    #[test]
    fn content_cells_skips_header() {
        let grid = GridLayoutBuilder::grid_n_m(2, 3, 0.0)
            .container(8.0, 6.0)
            .build();
        let cells: Vec<CellRect> = grid.content_cells(1).collect();
        assert_eq!(cells.len(), 4); // 2 cols × 2 remaining rows
    }

    // ── Existing ─────────────────────────────────────────────────────────

    #[test]
    fn to_position_props_roundtrip() {
        let r = CellRect { x: 1.5, y: 2.0, w: 4.0, h: 1.25 };
        let pp = r.to_position_props();
        assert!(matches!(pp.x, Some(Coord::Inches(v)) if approx(v, 1.5)));
        assert!(matches!(pp.y, Some(Coord::Inches(v)) if approx(v, 2.0)));
        assert!(matches!(pp.w, Some(Coord::Inches(v)) if approx(v, 4.0)));
        assert!(matches!(pp.h, Some(Coord::Inches(v)) if approx(v, 1.25)));
    }
}
