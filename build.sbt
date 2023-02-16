ThisBuild / scalaVersion := "3.2.2"
ThisBuild / version := "1.1.0"
ThisBuild / organization := "com.stulsoft"
ThisBuild / organizationName := "stulsoft"

lazy val poiVersion = "5.2.3"

lazy val root = (project in file("."))
  .settings(
    name := "ys-scala-poi",
    libraryDependencies += "com.typesafe.scala-logging" %% "scala-logging" % "3.9.5",
    libraryDependencies += "ch.qos.logback" % "logback-classic" % "1.4.5",
    libraryDependencies += "org.apache.logging.log4j" % "log4j-to-slf4j" % "2.19.0",

    libraryDependencies += "org.apache.poi" % "poi" % poiVersion,
    libraryDependencies += "org.apache.poi" % "poi-ooxml" % poiVersion,

    libraryDependencies += "org.scalatest" %% "scalatest" % "3.2.15" % Test
  )