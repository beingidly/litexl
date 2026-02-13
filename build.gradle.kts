plugins {
    java
    `java-library`
    jacoco
    alias(libs.plugins.graalvm.native)
    alias(libs.plugins.spotbugs)
    checkstyle
    alias(libs.plugins.maven.publish)
}

group = "com.beingidly"
version = "0.1.4"

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(libs.versions.java.get().toInt())
    }
    withJavadocJar()
    withSourcesJar()
}

repositories {
    mavenCentral()
}

dependencies {
    compileOnly(libs.jspecify)

    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)

    // Apache POI for cross-validation tests
    testImplementation(libs.poi.ooxml)
}

tasks.test {
    useJUnitPlatform()
    finalizedBy(tasks.jacocoTestReport)
}

jacoco {
    toolVersion = libs.versions.jacoco.get()
}

tasks.jacocoTestReport {
    dependsOn(tasks.test)
    reports {
        xml.required = true
        html.required = true
    }
}

tasks.jacocoTestCoverageVerification {
    violationRules {
        rule {
            limit {
                minimum = "0.90".toBigDecimal()
            }
        }
        rule {
            limit {
                counter = "BRANCH"
                minimum = "0.85".toBigDecimal()
            }
        }
    }
}

graalvmNative {
    binaries.all {
        buildArgs.add("--no-fallback")
        buildArgs.add("-H:+ReportExceptionStackTraces")
    }
    testSupport = true
}

// SpotBugs configuration
// Note: SpotBugs may not support the latest Java versions locally.
// CI runs with Java 21 for compatibility.
spotbugs {
    ignoreFailures = JavaVersion.current().majorVersion.toInt() > 21
    showStackTraces = true
    showProgress = true
    excludeFilter = file("config/spotbugs/exclude.xml")
}

tasks.withType<com.github.spotbugs.snom.SpotBugsTask>().configureEach {
    reports.create("html") { required = true }
    reports.create("xml") { required = true }
}

// Checkstyle configuration
checkstyle {
    toolVersion = libs.versions.checkstyle.get()
    configFile = file("config/checkstyle/checkstyle.xml")
    // Set to false once all violations are fixed
    isIgnoreFailures = true
}

tasks.withType<Checkstyle>().configureEach {
    reports {
        xml.required = true
        html.required = true
    }
}

// Maven Central publishing via vanniktech plugin
mavenPublishing {
    configure(com.vanniktech.maven.publish.JavaLibrary(
        javadocJar = com.vanniktech.maven.publish.JavadocJar.Javadoc(),
        sourcesJar = com.vanniktech.maven.publish.SourcesJar.Sources()
    ))
    publishToMavenCentral()
    signAllPublications()

    pom {
        name = "LiteXL"
        description = "Lightweight, zero-dependency Excel (XLSX) library for Java"
        url = "https://github.com/beingidly/litexl"
        inceptionYear = "2024"

        licenses {
            license {
                name = "The Apache License, Version 2.0"
                url = "https://www.apache.org/licenses/LICENSE-2.0.txt"
            }
        }

        developers {
            developer {
                id = "hssong"
                name = "Hyeonsik Song"
                email = "hssong@@beingidly.com"
            }
        }

        scm {
            connection = "scm:git:git://github.com/beingidly/litexl.git"
            developerConnection = "scm:git:ssh://github.com/beingidly/litexl.git"
            url = "https://github.com/beingidly/litexl"
        }
    }
}
