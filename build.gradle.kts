import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

plugins {
	kotlin("jvm") version "1.9.21"
}

java {
	sourceCompatibility = JavaVersion.VERSION_17
}

repositories {
	mavenCentral()
}

dependencies {
	//tomcat
	implementation("org.apache.tomcat.embed:tomcat-embed-core:10.1.17")

	//poi
	implementation("org.apache.poi:poi:4.1.2")
	implementation("org.apache.poi:poi-ooxml:4.1.2")
	implementation("org.apache.commons:commons-compress:1.21")

	//empty generator
	implementation("io.github.Yaklede:empty-object-generator:0.2.2")

	//kotlin
	implementation("org.jetbrains.kotlin:kotlin-reflect")

	///test
	implementation("org.junit.jupiter:junit-jupiter-api:5.7.2")
	testImplementation("org.junit.jupiter:junit-jupiter-engine:5.7.2")
	testImplementation("org.assertj:assertj-core:3.22.0")
}

tasks.withType<KotlinCompile> {
	kotlinOptions {
		freeCompilerArgs += "-Xjsr305=strict"
		jvmTarget = "17"
	}
}

tasks.withType<Test> {
	useJUnitPlatform()
}

