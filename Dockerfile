# 使用 OpenJDK 作为基础镜像
FROM openjdk:17-jre-slim

# 将本地的 JAR 包复制到容器中
COPY word-clear-0.0.1-SNAPSHOT /app/app.jar

# 设置工作目录
WORKDIR /app

# 运行 JAR 包
ENTRYPOINT ["java", "-jar", "word-clear-0.0.1-SNAPSHOT"]
