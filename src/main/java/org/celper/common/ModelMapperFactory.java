package org.celper.common;

import org.modelmapper.ModelMapper;
import org.modelmapper.config.Configuration;
import org.modelmapper.convention.MatchingStrategies;

public final class ModelMapperFactory {
    private ModelMapperFactory() {
        throw new IllegalStateException("Factory class");
    }

    public static ModelMapper defaultModelMapper() {
        ModelMapper modelMapper = new ModelMapper();
        modelMapper.getConfiguration().setMatchingStrategy(MatchingStrategies.STRICT);
        modelMapper.getConfiguration().setFieldMatchingEnabled(true);
        modelMapper.getConfiguration().setFieldAccessLevel(Configuration.AccessLevel.PRIVATE);
        addConverter(modelMapper);
        return modelMapper;
    }
    private static void addConverter(ModelMapper modelMapper) {
        modelMapper.addConverter(new ModelMapperConverters.DoubleToDate());
        modelMapper.addConverter(new ModelMapperConverters.DoubleToLocalDateTime());
        modelMapper.addConverter(new ModelMapperConverters.DoubleToLocalDate());
        modelMapper.addConverter(new ModelMapperConverters.DoubleToLocalTime());
        modelMapper.addConverter(new ModelMapperConverters.StringToDouble());
        modelMapper.addConverter(new ModelMapperConverters.StringToBoolean());
    }
}

