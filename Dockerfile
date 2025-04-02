# Use an official Maven image to build the app
FROM maven:3.9.4-eclipse-temurin-21 as builder
WORKDIR /app
COPY . .
RUN mvn clean install -DskipTests

# Use a minimal JDK runtime for running the app
FROM eclipse-temurin:21-jdk-alpine
WORKDIR /app
COPY --from=builder /app/target/*.jar app.jar

# Run the application
ENTRYPOINT ["java", "-jar", "app.jar"]
