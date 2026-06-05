/*
 * Copyright 2026 the poi-annotation-excel authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package io.github.dhsolo.common;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.function.Consumer;

/**
 * Reflection utilities used internally by the Excel framework to access fields and
 * methods across class hierarchies without requiring callers to handle checked exceptions.
 *
 * <p>All search methods walk the full superclass chain up to (but not including)
 * {@link Object}, making them suitable for working with domain objects that use
 * inheritance.  Accessibility is forced via {@link java.lang.reflect.AccessibleObject#setAccessible}
 * so that private and package-private members are reachable.
 *
 * <p>This class is {@code final} and has a private constructor; it is not meant to
 * be instantiated or subclassed.
 *
 * @author dhsolo
 * @since 1.0
 */
public final class Reflect {

    private Reflect() {}

    /**
     * Searches for a declared method by name and optional parameter types, walking the
     * superclass chain of {@code clazz} upward until a match is found or {@link Object}
     * is reached.
     *
     * <p>The returned method has its accessibility forced to {@code true}.
     *
     * @param clazz      the class to start the search from; must not be {@code null}
     * @param name       the simple name of the method to find
     * @param paramTypes the method's parameter types in declaration order; pass an empty
     *                   array or omit to search for a no-arg method
     * @return the first matching {@link Method} found in the hierarchy, or {@code null}
     *         if no such method exists
     */
    public static Method findMethod(Class<?> clazz, String name, Class<?>... paramTypes) {
        for (Class<?> c = clazz; c != null && c != Object.class; c = c.getSuperclass()) {
            try {
                Method m = paramTypes.length == 0
                        ? c.getDeclaredMethod(name)
                        : c.getDeclaredMethod(name, paramTypes);
                m.setAccessible(true);
                return m;
            } catch (NoSuchMethodException ignored) {}
        }
        return null;
    }

    /**
     * Searches for a declared field by name, walking the superclass chain of {@code clazz}
     * upward until a match is found or {@link Object} is reached.
     *
     * <p>The returned field has its accessibility forced to {@code true}.
     *
     * @param clazz the class to start the search from; must not be {@code null}
     * @param name  the simple name of the field to find
     * @return the first matching {@link Field} found in the hierarchy, or {@code null}
     *         if no such field exists
     */
    public static Field findField(Class<?> clazz, String name) {
        for (Class<?> c = clazz; c != null && c != Object.class; c = c.getSuperclass()) {
            try {
                Field f = c.getDeclaredField(name);
                f.setAccessible(true);
                return f;
            } catch (NoSuchFieldException ignored) {}
        }
        return null;
    }

    /**
     * Reads the value of a field from the given target object, bypassing access controls.
     *
     * @param field  the field to read; must not be {@code null}
     * @param target the object whose field value is to be retrieved; use {@code null}
     *               for static fields
     * @return the current value of the field on {@code target}
     * @throws RuntimeException wrapping {@link IllegalAccessException} if the field
     *                          value cannot be read
     */
    public static Object getField(Field field, Object target) {
        try {
            field.setAccessible(true);
            return field.get(target);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Invokes a method on the given target object with the supplied arguments, bypassing
     * access controls.
     *
     * @param method the method to invoke; must not be {@code null}
     * @param target the object on which to invoke the method; use {@code null} for
     *               static methods
     * @param args   the arguments to pass to the method; may be empty
     * @return the value returned by the method invocation, or {@code null} for
     *         {@code void} methods
     * @throws RuntimeException wrapping any {@link Exception} thrown during invocation,
     *                          including {@link java.lang.reflect.InvocationTargetException}
     */
    public static Object invokeMethod(Method method, Object target, Object... args) {
        try {
            method.setAccessible(true);
            return method.invoke(target, args);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Iterates over every declared field in {@code clazz} and all of its superclasses
     * (up to but not including {@link Object}), invoking {@code callback} for each field.
     *
     * <p>Fields are visited in declaration order within each class, starting from
     * {@code clazz} and ascending toward the root of the hierarchy.
     *
     * @param clazz    the class whose field hierarchy is to be traversed; must not be
     *                 {@code null}
     * @param callback a {@link Consumer} that receives each {@link Field}; must not be
     *                 {@code null}
     */
    public static void doWithFields(Class<?> clazz, Consumer<Field> callback) {
        for (Class<?> c = clazz; c != null && c != Object.class; c = c.getSuperclass()) {
            for (Field f : c.getDeclaredFields()) {
                callback.accept(f);
            }
        }
    }

    /**
     * Tests whether a string is non-null and contains at least one non-whitespace character.
     *
     * @param s the string to test, may be {@code null}
     * @return {@code true} if {@code s} has meaningful text content; {@code false} otherwise
     */
    public static boolean hasText(String s) {
        return s != null && !s.isBlank();
    }

    /**
     * Populates a new instance of {@code clazz} by mapping each entry in {@code map} to
     * the field whose name matches the entry's key.
     *
     * <p>The target class must expose a public no-arg constructor.  Keys in {@code map}
     * that do not correspond to a declared field in the class hierarchy are silently
     * ignored.  {@code null} map values are also skipped.
     *
     * @param <T>   the target bean type
     * @param map   a map of field names to their desired values; must not be {@code null}
     * @param clazz the class to instantiate and populate; must not be {@code null}
     * @return a new instance of {@code clazz} with fields set according to {@code map}
     * @throws RuntimeException if the no-arg constructor is absent or inaccessible, or
     *                          if a field assignment fails
     */
    public static <T> T mapToBean(java.util.Map<String, Object> map, Class<T> clazz) {
        try {
            T obj = clazz.getDeclaredConstructor().newInstance();
            for (java.util.Map.Entry<String, Object> entry : map.entrySet()) {
                Field f = findField(clazz, entry.getKey());
                if (f != null && entry.getValue() != null) {
                    f.set(obj, entry.getValue());
                }
            }
            return obj;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Produces a deep copy of the given {@link Serializable} object via Java serialization.
     *
     * <p>The object and every reachable object it references must be serializable.  This
     * method is primarily intended for cloning small Excel model objects; it is not
     * suitable for large or performance-critical data graphs.
     *
     * @param <T> the type of the object to clone; must extend {@link Serializable}
     * @param obj the object to clone; must not be {@code null}
     * @return a deep copy of {@code obj}
     * @throws RuntimeException wrapping {@link IOException} or
     *                          {@link ClassNotFoundException} if serialization or
     *                          deserialization fails
     */
    @SuppressWarnings("unchecked")
    public static <T extends Serializable> T clone(T obj) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            new ObjectOutputStream(baos).writeObject(obj);
            return (T) new ObjectInputStream(new ByteArrayInputStream(baos.toByteArray())).readObject();
        } catch (IOException | ClassNotFoundException e) {
            throw new RuntimeException(e);
        }
    }
}
