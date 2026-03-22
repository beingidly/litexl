plugins {
    java
    `java-library`
}

group = "com.beingidly"
version = "0.1.9"

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(libs.versions.java.get().toInt())
    }
}

repositories {
    mavenCentral()
}

dependencies {
    api(project(":"))

    compileOnly(libs.jspecify)

    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)

    // Apache POI for cross-validation tests
    testImplementation(libs.poi.ooxml) {
        exclude(group = "org.apache.poi", module = "poi-ooxml-lite")
    }
    testImplementation("org.apache.poi:poi-ooxml-full:${libs.versions.poi.get()}")
}

tasks.test {
    useJUnitPlatform()
}
