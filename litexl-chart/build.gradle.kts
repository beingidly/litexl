plugins {
    java
    `java-library`
    alias(libs.plugins.maven.publish)
}

group = "com.beingidly"
version = "0.1.10"

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(libs.versions.java.get().toInt())
    }
    withSourcesJar()
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

// Maven Central publishing via vanniktech plugin
mavenPublishing {
    configure(com.vanniktech.maven.publish.JavaLibrary(
        javadocJar = com.vanniktech.maven.publish.JavadocJar.Javadoc(),
        sourcesJar = com.vanniktech.maven.publish.SourcesJar.Sources()
    ))
    publishToMavenCentral()
    signAllPublications()

    pom {
        name = "LiteXL Chart"
        description = "Chart and graph support for LiteXL, a lightweight Excel (XLSX) library for Java"
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
