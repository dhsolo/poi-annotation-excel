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
package io.github.dhsolo.poi.excel;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.lang.reflect.Proxy;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Guards {@link ExcelModel#copy()} (used for {@code mergeCellIndex > 1} column clones): every
 * non-static field must be carried over. Each field is assigned a distinct non-default value and
 * the copy is compared field-by-field, so reverting to a hand-maintained copy that forgets a field
 * fails here instead of silently dropping configuration from a merged column.
 */
class ExcelModelCopyTest {

    @Test
    void copyPreservesEveryField() throws Exception {
        ExcelModel original = new ExcelModel("base");
        int seed = 1;
        for (Field f : ExcelModel.class.getDeclaredFields()) {
            if (Modifier.isStatic(f.getModifiers())) continue;
            Object sentinel = sampleValue(f.getType(), seed++);
            if (sentinel == null) continue; // type we cannot build a distinct value for
            f.setAccessible(true);
            f.set(original, sentinel);
        }

        ExcelModel copy = original.copy();

        for (Field f : ExcelModel.class.getDeclaredFields()) {
            if (Modifier.isStatic(f.getModifiers())) continue;
            f.setAccessible(true);
            assertThat(f.get(copy)).as("field '%s' must be carried over by copy()", f.getName())
                    .isEqualTo(f.get(original));
        }
    }

    /** A distinct, non-default value for the given field type, or {@code null} if none can be built. */
    private static Object sampleValue(Class<?> type, int seed) {
        if (type == boolean.class || type == Boolean.class) return Boolean.TRUE;
        if (type == int.class || type == Integer.class) return seed;
        if (type == long.class || type == Long.class) return (long) seed;
        if (type == short.class || type == Short.class) return (short) seed;
        if (type == byte.class || type == Byte.class) return (byte) seed;
        if (type == double.class || type == Double.class) return (double) seed;
        if (type == float.class || type == Float.class) return (float) seed;
        if (type == char.class || type == Character.class) return (char) ('a' + (seed % 26));
        if (type == String.class) return "v" + seed;
        if (type == Map.class) {
            Map<Object, Object> m = new HashMap<>();
            m.put("k" + seed, seed);
            return m;
        }
        if (type == Function.class) return (Function<Object, Object>) o -> null;
        if (type == ExcelModel.class) return new ExcelModel("parent" + seed);
        if (type.isInterface()) {
            // Shallow copy shares the reference, so any non-null instance works as a sentinel.
            // Handle the Object methods so the field-by-field equality assertion (which calls
            // equals/hashCode on this proxy) works via identity.
            return Proxy.newProxyInstance(type.getClassLoader(), new Class<?>[]{type},
                    (proxy, method, args) -> switch (method.getName()) {
                        case "equals" -> proxy == args[0];
                        case "hashCode" -> System.identityHashCode(proxy);
                        case "toString" -> "sentinel";
                        default -> null;
                    });
        }
        try {
            var ctor = type.getDeclaredConstructor();
            ctor.setAccessible(true);
            return ctor.newInstance();
        } catch (Exception cannotBuild) {
            return null;
        }
    }
}
