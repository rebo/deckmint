pub mod chart_xml;
pub mod pres_xml;
pub mod rels_xml;
pub mod slide_xml;
pub mod table_xml;

/// Carriage-return + line-feed, as used in OOXML files
pub const CRLF: &str = "\r\n";

/// A simple macro to append formatted strings to a String buffer.
/// Equivalent to `write!(buf, ...).expect(...)` but cleaner at call sites.
#[allow(unused_macros)]
macro_rules! xml {
    ($buf:expr, $($arg:tt)*) => {
        {
            use std::fmt::Write;
            write!($buf, $($arg)*).expect("write to String is infallible")
        }
    };
}

#[allow(unused_imports)]
pub(crate) use xml;
