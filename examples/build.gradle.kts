plugins {
    java
    application
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
    implementation(project(":"))
}

// Default main class (can be overridden with -PmainClass=...)
application {
    mainClass = project.findProperty("mainClass") as String?
        ?: "com.beingidly.litexl.examples.core.Ex01_CreateWorkbook"
}

// Create run tasks for each example category
val categories = listOf("core", "mapper", "security", "advanced")

tasks.register("listExamples") {
    group = "examples"
    description = "List all available examples"
    doLast {
        val examplesDir = file("src/main/java/com/beingidly/litexl/examples")
        categories.forEach { category ->
            val categoryDir = examplesDir.resolve(category)
            if (categoryDir.exists()) {
                println("\n=== $category ===")
                categoryDir.listFiles()
                    ?.filter { it.name.endsWith(".java") }
                    ?.sortedBy { it.name }
                    ?.forEach { file ->
                        val className = file.nameWithoutExtension
                        println("  ./gradlew :examples:run -PmainClass=com.beingidly.litexl.examples.$category.$className")
                    }
            }
        }
    }
}
