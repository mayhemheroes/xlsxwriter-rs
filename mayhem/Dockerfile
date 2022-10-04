FROM ghcr.io/evanrichter/cargo-fuzz as builder

ADD . /xlsxwriter
WORKDIR /xlsxwriter/fuzz
RUN cargo +nightly fuzz build 

FROM debian:bookworm
COPY --from=builder /xlsxwriter/fuzz/target/x86_64-unknown-linux-gnu/release/xlsxwriter-fuzz /