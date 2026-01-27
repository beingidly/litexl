plugins {
    java
    application
    id("me.champeau.jmh") version "0.7.2"
    id("org.graalvm.buildtools.native") version "0.10.4"
}

application {
    mainClass = "com.beingidly.litexl.benchmark.QuickCompare"
}

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(25)
    }
}

repositories {
    mavenCentral()
}

dependencies {
    // Main library
    implementation(project(":"))

    // Apache POI for comparison (JVM only)
    implementation("org.apache.poi:poi-ooxml:5.3.0")

    // JMH
    jmh("org.openjdk.jmh:jmh-core:1.37")
    jmh("org.openjdk.jmh:jmh-generator-annprocess:1.37")
}

tasks.test {
    useJUnitPlatform()
}

jmh {
    warmupIterations = 1
    iterations = 2
    fork = 1
    resultsFile = project.file("build/reports/jmh/results.json")
    resultFormat = "JSON"

    // GC profiling
    profilers = listOf("gc")
}

tasks.named("jmhJar") {
    // Ensure main project is built first
    dependsOn(":jar")
}

// Native image configuration for NativeBenchmark
graalvmNative {
    binaries {
        named("main") {
            imageName = "litexl-native-benchmark"
            mainClass = "com.beingidly.litexl.benchmark.NativeBenchmark"
            buildArgs.add("--no-fallback")
            buildArgs.add("-H:+ReportExceptionStackTraces")
        }
    }
}

// Task to run JVM benchmark
tasks.register<JavaExec>("jvmBenchmark") {
    group = "benchmark"
    description = "Run litexl benchmark on JVM"
    classpath = sourceSets["main"].runtimeClasspath
    mainClass = "com.beingidly.litexl.benchmark.NativeBenchmark"
    args = listOf("3", "10000", "30")
}

// Task to run complex comparison
tasks.register<JavaExec>("complexCompare") {
    group = "benchmark"
    description = "Run complex benchmark comparison (litexl vs POI, all features)"
    classpath = sourceSets["main"].runtimeClasspath
    mainClass = "com.beingidly.litexl.benchmark.ComplexCompare"
    jvmArgs = listOf("-Xmx1g")
}

// Task to run JMH complex benchmark
tasks.register<JavaExec>("complexBenchmark") {
    group = "benchmark"
    description = "Run JMH complex benchmark"
    dependsOn("jmhJar")
    classpath = files(tasks.named("jmhJar").map { (it as Jar).archiveFile })
    mainClass = "org.openjdk.jmh.Main"
    args = listOf(
        "ComplexBenchmark",
        "-rf", "json",
        "-rff", "build/reports/jmh/complex-results.json",
        "-prof", "gc"
    )
    jvmArgs = listOf("-Xmx1g")
}
