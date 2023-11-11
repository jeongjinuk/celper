package org.celper.util;

import org.celper.core.structure.Structure;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * The type Structure registrator.
 */
public final class StructureRegistrator {
    private final Map<Class<?>, List<Structure>> structureMap;

    /**
     * Instantiates a new Structure registrator.
     */
    public StructureRegistrator() {
        this.structureMap = new HashMap<>();
    }

    /**
     * Gets or default.
     *
     * @param clazz the clazz
     * @return the or default
     */
    public List<Structure> getOrDefault(Class<?> clazz) {
        return structureMap.get(clazz) == null ? createStructures(clazz) : structureMap.get(clazz);
    }

    /**
     * Add.
     *
     * @param clazz the clazz
     */
    public void add(Class<?> clazz) {
        structureMap.put(clazz, createStructures(clazz));
    }

    /**
     * Create structures list.
     *
     * @param clazz the clazz
     * @return the list
     */
    public List<Structure> createStructures(Class<?> clazz) {
        AtomicInteger atomicInteger = new AtomicInteger(0);
        return ReflectionUtils.getDeclaredFields(clazz)
                .stream()
                .filter(Structure :: existsColumnAnnotation)
                .map(field -> new Structure(clazz, field, atomicInteger.getAndIncrement()))
                .collect(Collectors.toList());
    }
}
